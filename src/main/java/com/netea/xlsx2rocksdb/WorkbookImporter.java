package com.netea.xlsx2rocksdb;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.rocksdb.RocksDBException;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;
import javax.xml.parsers.ParserConfigurationException;

final class WorkbookImporter {
    private static final int MIN_PROGRESS_INTERVAL = 10_000;

    List<SheetResult> importWorkbook(
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
        try (OPCPackage packageFile = OPCPackage.open(input.toFile(), PackageAccess.READ)) {
            XSSFReader reader = new XSSFReader(packageFile);
            reader.setUseReadOnlySharedStringsTable(true);
            Styles styles = reader.getStylesTable();
            SharedStrings strings = reader.getSharedStringsTable();
            DataFormatter formatter = new DataFormatter(Locale.ROOT);

            progress.accept("Scanning workbook structure.");
            AnalysisPass analysisPass = analyzeSheets(
                    reader, styles, strings, formatter, selectedSheets, configuredHeaderRow, progress);
            validateSelectedSheets(analysisPass.sheetNames(), selectedSheets);
            progress.accept("Workbook opened: " + analysisPass.sheetNames().size() + " sheets.");

            return importSheets(reader, styles, strings, formatter, input, selectedSheets,
                    source, columnSelectors, keySelectors, analysisPass.analyses(), store, progress, warning);
        } catch (ImportFailure failure) {
            throw failure.cause();
        } catch (OpenXML4JException | SAXException exception) {
            throw new IOException("Could not stream XLSX workbook: " + exception.getMessage(), exception);
        }
    }

    private AnalysisPass analyzeSheets(
            XSSFReader reader,
            Styles styles,
            SharedStrings strings,
            DataFormatter formatter,
            Set<String> selectedSheets,
            Integer configuredHeaderRow,
            Consumer<String> progress
    ) throws IOException, OpenXML4JException, SAXException {
        List<String> sheetNames = new ArrayList<>();
        Map<String, SheetAnalysis> analyses = new LinkedHashMap<>();
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();

        while (sheets.hasNext()) {
            try (InputStream stream = sheets.next()) {
                String sheetName = sheets.getSheetName();
                sheetNames.add(sheetName);
                if (!selectedSheets.isEmpty() && !selectedSheets.contains(sheetName)) {
                    continue;
                }
                progress.accept("Sheet '" + sheetName + "': analyzing header and columns.");
                AnalysisHandler handler = new AnalysisHandler(configuredHeaderRow);
                parseSheet(stream, styles, strings, formatter, handler);
                analyses.put(sheetName, handler.result());
            }
        }
        return new AnalysisPass(List.copyOf(sheetNames), analyses);
    }

    private List<SheetResult> importSheets(
            XSSFReader reader,
            Styles styles,
            SharedStrings strings,
            DataFormatter formatter,
            Path input,
            Set<String> selectedSheets,
            String source,
            List<String> columnSelectors,
            List<String> keySelectors,
            Map<String, SheetAnalysis> analyses,
            RocksDbStore store,
            Consumer<String> progress,
            Consumer<String> warning
    ) throws IOException, OpenXML4JException, SAXException, RocksDBException {
        List<SheetResult> results = new ArrayList<>();
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();

        while (sheets.hasNext()) {
            try (InputStream stream = sheets.next()) {
                String sheetName = sheets.getSheetName();
                if (!selectedSheets.isEmpty() && !selectedSheets.contains(sheetName)) {
                    continue;
                }
                results.add(importSheet(stream, sheetName, analyses.get(sheetName), styles, strings,
                        formatter, source, columnSelectors, keySelectors, store, progress, warning));
            }
        }
        return results;
    }

    private SheetResult importSheet(
            InputStream stream,
            String sheetName,
            SheetAnalysis analysis,
            Styles styles,
            SharedStrings strings,
            DataFormatter formatter,
            String source,
            List<String> selectors,
            List<String> keySelectors,
            RocksDbStore store,
            Consumer<String> progress,
            Consumer<String> warning
    ) throws RocksDBException, IOException, SAXException {
        if (analysis.headerIndex() == null) {
            if (!selectors.isEmpty() || !keySelectors.isEmpty()) {
                throw new IllegalArgumentException("Cannot select columns from empty sheet '" + sheetName + "'");
            }
            progress.accept("Sheet '" + sheetName + "': empty; writing metadata only.");
            writeEmptySheet(store, sheetName, source, null);
            return new SheetResult(sheetName, 0, 0, null);
        }

        if (analysis.width() == 0) {
            if (!selectors.isEmpty() || !keySelectors.isEmpty()) {
                throw new IllegalArgumentException("Cannot select columns from empty sheet '" + sheetName + "'");
            }
            progress.accept("Sheet '" + sheetName + "': no data columns; writing metadata only.");
            writeEmptySheet(store, sheetName, source, analysis.headerIndex() + 1);
            return new SheetResult(sheetName, 0, 0, analysis.headerIndex() + 1);
        }

        List<String> allColumnNames = createColumnNames(analysis.headerValues(), analysis.width());
        List<Integer> selectedIndexes = resolveColumns(selectors, allColumnNames, sheetName);
        List<String> selectedNames = selectedIndexes.stream().map(allColumnNames::get).toList();
        List<Integer> keyIndexes = resolveKeyColumns(keySelectors, allColumnNames, sheetName);
        List<String> keyNames = keyIndexes.stream().map(allColumnNames::get).toList();
        progress.accept("Sheet '" + sheetName + "': header row " + (analysis.headerIndex() + 1)
                + ", " + selectedNames.size() + " of " + allColumnNames.size() + " columns selected.");
        progress.accept("Sheet '" + sheetName + "': preparing RocksDB table.");

        RocksDbStore.SheetWriter writer = store.startSheet(
                sheetName, source, analysis.headerIndex() + 1, selectedNames, keyNames);
        ImportHandler handler = new ImportHandler(
                sheetName, analysis, selectedIndexes, selectedNames,
                keyIndexes, keyNames, writer, progress, warning);
        try {
            parseSheet(stream, styles, strings, formatter, handler);
            writer.finish();
        } finally {
            writer.close();
        }
        progress.accept("Sheet '" + sheetName + "': complete; " + handler.rowCount() + " rows imported.");
        return new SheetResult(sheetName, handler.rowCount(), selectedNames.size(), analysis.headerIndex() + 1);
    }

    private static void parseSheet(
            InputStream stream,
            Styles styles,
            SharedStrings strings,
            DataFormatter formatter,
            XSSFSheetXMLHandler.SheetContentsHandler contentsHandler
    ) throws SAXException, IOException {
        XMLReader parser;
        try {
            parser = XMLHelper.newXMLReader();
        } catch (ParserConfigurationException exception) {
            throw new SAXException("Could not configure secure XLSX parser", exception);
        }
        parser.setContentHandler(new XSSFSheetXMLHandler(
                styles, null, strings, contentsHandler, formatter, false));
        parser.parse(new InputSource(stream));
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

    static List<String> createColumnNames(Map<Integer, String> headerValues, int width) {
        List<String> names = new ArrayList<>(width);
        Map<String, Integer> occurrences = new HashMap<>();
        for (int index = 0; index < width; index++) {
            String candidate = headerValues.getOrDefault(index, "").trim();
            String baseName = candidate.isEmpty() ? "col" + (index + 1) : candidate;
            int occurrence = occurrences.merge(baseName, 1, Integer::sum);
            names.add(occurrence == 1 ? baseName : baseName + "_" + occurrence);
        }
        return names;
    }

    static List<Integer> resolveColumns(
            List<String> selectors, List<String> names, String sheetName) {
        return resolveColumns(selectors, names, sheetName, true, false);
    }

    static List<Integer> resolveKeyColumns(
            List<String> selectors, List<String> names, String sheetName) {
        return resolveColumns(selectors, names, sheetName, false, true);
    }

    private static List<Integer> resolveColumns(
            List<String> selectors,
            List<String> names,
            String sheetName,
            boolean defaultAll,
            boolean numericIsIndex
    ) {
        if (selectors.isEmpty()) {
            if (!defaultAll) {
                return List.of();
            }
            List<Integer> all = new ArrayList<>(names.size());
            for (int index = 0; index < names.size(); index++) {
                all.add(index);
            }
            return all;
        }

        Set<Integer> indexes = new LinkedHashSet<>();
        for (String rawSelector : selectors) {
            String selector = rawSelector.trim();
            if (selector.isEmpty()) {
                throw new IllegalArgumentException("Column selectors must not be empty");
            }
            int index = resolveColumn(selector, names, numericIsIndex);
            if (index < 0 || index >= names.size()) {
                throw new IllegalArgumentException(
                        "Column selector '" + selector + "' is out of range for sheet '" + sheetName + "'");
            }
            indexes.add(index);
        }
        return List.copyOf(indexes);
    }

    private static int resolveColumn(String selector, List<String> names, boolean numericIsIndex) {
        if (selector.startsWith("name:")) {
            return requireNamedColumn(selector.substring("name:".length()), names);
        }
        if (selector.startsWith("index:")) {
            return parseOneBasedIndex(selector.substring("index:".length()), selector);
        }
        if (selector.startsWith("letter:")) {
            return parseColumnLetters(selector.substring("letter:".length()), selector);
        }

        if (numericIsIndex && selector.chars().allMatch(Character::isDigit)) {
            return parseOneBasedIndex(selector, selector);
        }
        int namedIndex = names.indexOf(selector);
        if (namedIndex >= 0) {
            return namedIndex;
        }
        if (selector.chars().allMatch(Character::isDigit)) {
            return parseOneBasedIndex(selector, selector);
        }
        if (selector.chars().allMatch(Character::isLetter)) {
            return parseColumnLetters(selector, selector);
        }
        throw new IllegalArgumentException("Unknown column '" + selector + "'; available columns: " + names);
    }

    private static int requireNamedColumn(String name, List<String> names) {
        int index = names.indexOf(name);
        if (index < 0) {
            throw new IllegalArgumentException("Unknown column '" + name + "'; available columns: " + names);
        }
        return index;
    }

    private static int parseOneBasedIndex(String value, String selector) {
        try {
            int index = Integer.parseInt(value);
            if (index < 1) {
                throw new NumberFormatException();
            }
            return index - 1;
        } catch (NumberFormatException exception) {
            throw new IllegalArgumentException("Invalid one-based column selector: '" + selector + "'");
        }
    }

    private static int parseColumnLetters(String value, String selector) {
        if (value.isEmpty()) {
            throw new IllegalArgumentException("Invalid Excel column selector: '" + selector + "'");
        }
        int result = 0;
        try {
            for (char character : value.toUpperCase(Locale.ROOT).toCharArray()) {
                if (character < 'A' || character > 'Z') {
                    throw new IllegalArgumentException("Invalid Excel column selector: '" + selector + "'");
                }
                result = Math.addExact(Math.multiplyExact(result, 26), character - 'A' + 1);
            }
        } catch (ArithmeticException exception) {
            throw new IllegalArgumentException("Invalid Excel column selector: '" + selector + "'");
        }
        return result - 1;
    }

    private static void validateSelectedSheets(List<String> sheetNames, Set<String> selectedSheets) {
        if (selectedSheets.isEmpty()) {
            return;
        }
        Set<String> available = new HashSet<>(sheetNames);
        Set<String> missing = new LinkedHashSet<>(selectedSheets);
        missing.removeAll(available);
        if (!missing.isEmpty()) {
            throw new IllegalArgumentException(
                    "Workbook does not contain sheet(s) " + missing + "; available: " + available);
        }
    }

    static int progressInterval(int totalSourceRows) {
        int tenPercent = (int) Math.min(Integer.MAX_VALUE, ((long) Math.max(0, totalSourceRows) + 9) / 10);
        return Math.max(MIN_PROGRESS_INTERVAL, tenPercent);
    }

    record SheetResult(String sheetName, int rowCount, int columnCount, Integer headerRow) {
    }

    private record AnalysisPass(List<String> sheetNames, Map<String, SheetAnalysis> analyses) {
    }

    private record SheetAnalysis(
            Integer headerIndex, Map<Integer, String> headerValues, int width, int lastRowIndex) {
    }

    private static final class AnalysisHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private final Integer configuredHeaderIndex;
        private Integer headerIndex;
        private Map<Integer, String> headerValues = Map.of();
        private int currentRow;
        private int currentWidth;
        private int nextColumn;
        private int width;
        private int lastRowIndex = -1;
        private boolean currentHasValue;
        private Map<Integer, String> currentValues = new LinkedHashMap<>();

        private AnalysisHandler(Integer configuredHeaderRow) {
            configuredHeaderIndex = configuredHeaderRow == null ? null : configuredHeaderRow - 1;
            headerIndex = configuredHeaderIndex;
        }

        @Override
        public void startRow(int rowNum) {
            currentRow = rowNum;
            currentWidth = 0;
            nextColumn = 0;
            currentHasValue = false;
            currentValues = new LinkedHashMap<>();
            lastRowIndex = Math.max(lastRowIndex, rowNum);
        }

        @Override
        public void endRow(int rowNum) {
            if (configuredHeaderIndex == null && headerIndex == null && currentHasValue) {
                headerIndex = rowNum;
                headerValues = Map.copyOf(currentValues);
            } else if (headerIndex != null && rowNum == headerIndex) {
                headerValues = Map.copyOf(currentValues);
            }
            if (headerIndex != null && rowNum >= headerIndex) {
                width = Math.max(width, currentWidth);
            }
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            int column = columnIndex(cellReference, nextColumn);
            nextColumn = column + 1;
            currentWidth = Math.max(currentWidth, column + 1);
            String value = formattedValue == null ? "" : formattedValue;
            currentValues.put(column, value);
            if (!value.isBlank()) {
                currentHasValue = true;
            }
        }

        SheetAnalysis result() {
            return new SheetAnalysis(headerIndex, Map.copyOf(headerValues), width, lastRowIndex);
        }
    }

    private static final class ImportHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
        private final String sheetName;
        private final SheetAnalysis analysis;
        private final List<Integer> selectedIndexes;
        private final List<String> selectedNames;
        private final List<Integer> keyIndexes;
        private final List<String> keyNames;
        private final RocksDbStore.SheetWriter writer;
        private final Consumer<String> progress;
        private final Consumer<String> warning;
        private final int totalSourceRows;
        private final int progressInterval;
        private Map<Integer, String> currentValues = new HashMap<>();
        private int nextColumn;
        private boolean currentHasValue;
        private int rowCount;
        private int nextProgress;

        private ImportHandler(
                String sheetName,
                SheetAnalysis analysis,
                List<Integer> selectedIndexes,
                List<String> selectedNames,
                List<Integer> keyIndexes,
                List<String> keyNames,
                RocksDbStore.SheetWriter writer,
                Consumer<String> progress,
                Consumer<String> warning
        ) {
            this.sheetName = sheetName;
            this.analysis = analysis;
            this.selectedIndexes = selectedIndexes;
            this.selectedNames = selectedNames;
            this.keyIndexes = keyIndexes;
            this.keyNames = keyNames;
            this.writer = writer;
            this.progress = progress;
            this.warning = warning;
            totalSourceRows = Math.max(0, analysis.lastRowIndex() - analysis.headerIndex());
            progressInterval = progressInterval(totalSourceRows);
            nextProgress = progressInterval;
        }

        @Override
        public void startRow(int rowNum) {
            currentValues = new HashMap<>();
            nextColumn = 0;
            currentHasValue = false;
        }

        @Override
        public void endRow(int rowNum) {
            if (rowNum > analysis.headerIndex() && currentHasValue) {
                Map<String, String> values = new LinkedHashMap<>();
                for (int index = 0; index < selectedIndexes.size(); index++) {
                    values.put(selectedNames.get(index), currentValues.getOrDefault(selectedIndexes.get(index), ""));
                }
                Map<String, String> keyValues = new LinkedHashMap<>();
                for (int index = 0; index < keyIndexes.size(); index++) {
                    keyValues.put(keyNames.get(index), currentValues.getOrDefault(keyIndexes.get(index), ""));
                }
                try {
                    writer.putRow(rowNum + 1, values, keyValues);
                } catch (RocksDBException exception) {
                    throw new ImportFailure(exception);
                }
                rowCount++;
            }

            int scannedRows = rowNum - analysis.headerIndex();
            if (scannedRows >= nextProgress && scannedRows < totalSourceRows) {
                int percentage = (int) ((long) scannedRows * 100 / totalSourceRows);
                progress.accept("Sheet '" + sheetName + "': " + scannedRows + "/"
                        + totalSourceRows + " source rows scanned (" + percentage + "%).");
                nextProgress = ((scannedRows / progressInterval) + 1) * progressInterval;
            }
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            int column = columnIndex(cellReference, nextColumn);
            nextColumn = column + 1;
            String value = formattedValue == null ? "" : formattedValue;
            currentValues.put(column, value);
            if (!value.isBlank()) {
                currentHasValue = true;
            }
        }

        int rowCount() {
            return rowCount;
        }
    }

    private static int columnIndex(String cellReference, int fallback) {
        return cellReference == null ? fallback : new CellReference(cellReference).getCol();
    }

    private static final class ImportFailure extends RuntimeException {
        private final RocksDBException cause;

        private ImportFailure(RocksDBException cause) {
            super(cause);
            this.cause = cause;
        }

        private RocksDBException cause() {
            return cause;
        }
    }
}
