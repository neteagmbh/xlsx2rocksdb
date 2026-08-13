package com.netea.xlsx2rocksdb;

import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;
import org.rocksdb.RocksDBException;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;
import java.util.stream.Stream;

final class FastExcelWorkbookImporter {
    List<WorkbookImporter.SheetResult> importWorkbook(
            Path input,
            String source,
            Set<String> selectedSheets,
            Integer configuredHeaderRow,
            List<String> columnSelectors,
            List<String> keySelectors,
            RocksDbStore store,
            Consumer<String> progress,
            Consumer<String> warning
    ) throws IOException, RocksDBException {
        try (ReadableWorkbook workbook = new ReadableWorkbook(input.toFile())) {
            List<Sheet> sheets = workbook.getSheets().toList();
            List<String> sheetNames = sheets.stream().map(Sheet::getName).toList();
            validateSelectedSheets(sheetNames, selectedSheets);
            progress.accept("Workbook opened: " + sheetNames.size() + " sheets.");

            List<WorkbookImporter.SheetResult> results = new ArrayList<>();
            for (Sheet sheet : sheets) {
                if (!selectedSheets.isEmpty() && !selectedSheets.contains(sheet.getName())) {
                    continue;
                }
                progress.accept("Sheet '" + sheet.getName() + "': analyzing header and columns.");
                SheetAnalysis analysis = analyzeSheet(sheet, configuredHeaderRow);
                results.add(importSheet(
                        sheet, analysis, source, columnSelectors, keySelectors, store, progress, warning));
            }
            return results;
        }
    }

    private static SheetAnalysis analyzeSheet(Sheet sheet, Integer configuredHeaderRow) throws IOException {
        Integer configuredHeaderIndex = configuredHeaderRow == null ? null : configuredHeaderRow - 1;
        Integer headerIndex = configuredHeaderIndex;
        Map<Integer, String> headerValues = Map.of();
        int width = 0;
        int lastRowIndex = -1;

        try (Stream<Row> rows = sheet.openStream()) {
            for (var iterator = rows.iterator(); iterator.hasNext();) {
                Row row = iterator.next();
                int rowIndex = row.getRowNum() - 1;
                lastRowIndex = Math.max(lastRowIndex, rowIndex);
                boolean hasValue = rowHasValue(row);
                if (configuredHeaderIndex == null && headerIndex == null && hasValue) {
                    headerIndex = rowIndex;
                    headerValues = values(row);
                } else if (headerIndex != null && rowIndex == headerIndex) {
                    headerValues = values(row);
                }
                if (headerIndex != null && rowIndex >= headerIndex) {
                    width = Math.max(width, row.getCellCount());
                }
            }
        }
        return new SheetAnalysis(headerIndex, headerValues, width, lastRowIndex);
    }

    private static WorkbookImporter.SheetResult importSheet(
            Sheet sheet,
            SheetAnalysis analysis,
            String source,
            List<String> columnSelectors,
            List<String> keySelectors,
            RocksDbStore store,
            Consumer<String> progress,
            Consumer<String> warning
    ) throws IOException, RocksDBException {
        String sheetName = sheet.getName();
        if (analysis.headerIndex() == null) {
            if (!columnSelectors.isEmpty() || !keySelectors.isEmpty()) {
                throw new IllegalArgumentException("Cannot select columns from empty sheet '" + sheetName + "'");
            }
            progress.accept("Sheet '" + sheetName + "': empty; writing metadata only.");
            writeEmptySheet(store, sheetName, source, null);
            return new WorkbookImporter.SheetResult(sheetName, 0, 0, null);
        }

        if (analysis.width() == 0) {
            if (!columnSelectors.isEmpty() || !keySelectors.isEmpty()) {
                throw new IllegalArgumentException("Cannot select columns from empty sheet '" + sheetName + "'");
            }
            progress.accept("Sheet '" + sheetName + "': no data columns; writing metadata only.");
            writeEmptySheet(store, sheetName, source, analysis.headerIndex() + 1);
            return new WorkbookImporter.SheetResult(sheetName, 0, 0, analysis.headerIndex() + 1);
        }

        List<String> allNames = WorkbookImporter.createColumnNames(analysis.headerValues(), analysis.width());
        List<Integer> selectedIndexes = WorkbookImporter.resolveColumns(columnSelectors, allNames, sheetName);
        List<String> selectedNames = selectedIndexes.stream().map(allNames::get).toList();
        List<Integer> keyIndexes = WorkbookImporter.resolveKeyColumns(keySelectors, allNames, sheetName);
        List<String> keyNames = keyIndexes.stream().map(allNames::get).toList();
        progress.accept("Sheet '" + sheetName + "': header row " + (analysis.headerIndex() + 1)
                + ", " + selectedNames.size() + " of " + allNames.size() + " columns selected.");
        progress.accept("Sheet '" + sheetName + "': preparing RocksDB table.");

        RocksDbStore.SheetWriter writer = store.startSheet(
                sheetName, source, analysis.headerIndex() + 1, selectedNames, keyNames);
        int rowCount = 0;
        int totalSourceRows = Math.max(0, analysis.lastRowIndex() - analysis.headerIndex());
        int interval = WorkbookImporter.progressInterval(totalSourceRows);
        int nextProgress = interval;
        try (Stream<Row> rows = sheet.openStream()) {
            for (var iterator = rows.iterator(); iterator.hasNext();) {
                Row row = iterator.next();
                int rowIndex = row.getRowNum() - 1;
                if (rowIndex > analysis.headerIndex() && rowHasValue(row)) {
                    Map<String, String> values = selectedValues(row, selectedIndexes, selectedNames);
                    Map<String, String> keyValues = selectedValues(row, keyIndexes, keyNames);
                    writer.putRow(rowIndex + 1, values, keyValues);
                    rowCount++;
                }
                int scannedRows = rowIndex - analysis.headerIndex();
                if (scannedRows >= nextProgress && scannedRows < totalSourceRows) {
                    int percentage = (int) ((long) scannedRows * 100 / totalSourceRows);
                    progress.accept("Sheet '" + sheetName + "': " + scannedRows + "/"
                            + totalSourceRows + " source rows scanned (" + percentage + "%).");
                    nextProgress = ((scannedRows / interval) + 1) * interval;
                }
            }
            writer.finish();
        } finally {
            writer.close();
        }
        progress.accept("Sheet '" + sheetName + "': complete; " + rowCount + " rows imported.");
        return new WorkbookImporter.SheetResult(
                sheetName, rowCount, selectedNames.size(), analysis.headerIndex() + 1);
    }

    private static Map<Integer, String> values(Row row) {
        Map<Integer, String> values = new LinkedHashMap<>();
        row.forEach(cell -> {
            if (cell != null) {
                values.put(cell.getColumnIndex(), cell.getText());
            }
        });
        return Map.copyOf(values);
    }

    private static boolean rowHasValue(Row row) {
        return row.stream().anyMatch(cell -> cell != null && !cell.getText().isBlank());
    }

    private static Map<String, String> selectedValues(
            Row row, List<Integer> indexes, List<String> names) {
        Map<String, String> values = new LinkedHashMap<>();
        for (int index = 0; index < indexes.size(); index++) {
            values.put(names.get(index), row.getCellText(indexes.get(index)));
        }
        return values;
    }

    private static void writeEmptySheet(
            RocksDbStore store, String sheetName, String source, Integer headerRow) throws RocksDBException {
        RocksDbStore.SheetWriter writer = store.startSheet(
                sheetName, source, headerRow, List.of(), List.of());
        try {
            writer.finish();
        } finally {
            writer.close();
        }
    }

    private static void validateSelectedSheets(List<String> sheetNames, Set<String> selectedSheets) {
        if (selectedSheets.isEmpty()) {
            return;
        }
        Set<String> missing = new LinkedHashSet<>(selectedSheets);
        missing.removeAll(new HashSet<>(sheetNames));
        if (!missing.isEmpty()) {
            throw new IllegalArgumentException(
                    "Workbook does not contain sheet(s) " + missing + "; available: " + sheetNames);
        }
    }

    private record SheetAnalysis(
            Integer headerIndex, Map<Integer, String> headerValues, int width, int lastRowIndex) {
    }
}
