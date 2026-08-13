package com.netea.xlsx2rocksdb;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.rocksdb.ColumnFamilyDescriptor;
import org.rocksdb.ColumnFamilyHandle;
import org.rocksdb.ColumnFamilyOptions;
import org.rocksdb.DBOptions;
import org.rocksdb.RocksDB;
import org.rocksdb.RocksDBException;
import picocli.CommandLine;

import java.io.IOException;
import java.io.ByteArrayInputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ImportCommandTest {
    private static final ObjectMapper JSON = new ObjectMapper();

    @TempDir
    Path temporaryDirectory;

    @Test
    void importsSelectedSheetUsingFirstNonEmptyHeaderAndColumnSubset() throws Exception {
        Path workbook = temporaryDirectory.resolve("input.xlsx");
        try (XSSFWorkbook xlsx = new XSSFWorkbook()) {
            var customers = xlsx.createSheet("Customers");
            customers.createRow(0);
            var header = customers.createRow(2);
            header.createCell(0).setCellValue("id");
            header.createCell(1).setCellValue("name");
            header.createCell(2).setCellValue("active");
            var data = customers.createRow(3);
            data.createCell(0).setCellValue(42);
            data.createCell(1).setCellValue("Ada");
            data.createCell(2).setCellValue(true);
            var duplicate = customers.createRow(4);
            duplicate.createCell(0).setCellValue(42);
            duplicate.createCell(1).setCellValue("Grace");
            duplicate.createCell(2).setCellValue(false);

            var ignored = xlsx.createSheet("Ignored");
            ignored.createRow(0).createCell(0).setCellValue("value");
            write(xlsx, workbook);
        }

        Path database = temporaryDirectory.resolve("database");
        StringWriter output = new StringWriter();
        StringWriter errors = new StringWriter();
        CommandLine command = new CommandLine(new ImportCommand());
        command.setOut(new PrintWriter(output, true));
        command.setErr(new PrintWriter(errors, true));
        int exitCode = command.execute(
                workbook.toString(), "--db", database.toString(),
                "--engine", "fastexcel", "--sheets", "Customers",
                "--columns", "name,index:3", "--key", "1,2",
                "--batch-size", "1", "--verbose");

        assertEquals(0, exitCode);
        assertTrue(output.toString().contains("Opening RocksDB:"));
        assertTrue(output.toString().contains("Excel engine: fastexcel"));
        assertTrue(output.toString().contains("Sheet 'Customers': analyzing header and columns."));
        assertTrue(output.toString().contains("Sheet 'Customers': complete; 2 rows imported."));
        assertEquals("", errors.toString());
        assertTrue(familyNames(database).contains("Customers"));
        assertFalse(familyNames(database).contains("Ignored"));
        try (OpenedFamily opened = openFamily(database, "Customers")) {
            JsonNode metadata = JSON.readTree(opened.database().get(opened.handle(), RocksDbStore.METADATA_KEY));
            assertEquals(3, metadata.get("headerRow").asInt());
            assertEquals(List.of("name", "active"), JSON.convertValue(metadata.get("columns"), List.class));
            assertEquals("id", metadata.get("keyColumns").get("id").asText());
            assertEquals("name", metadata.get("keyColumns").get("name").asText());
            assertEquals(2, metadata.get("rowCount").asInt());
            JsonNode availableKeys = JSON.readTree(
                    opened.database().get(opened.handle(), RocksDbStore.KEYS_KEY));
            assertEquals("key:id:", availableKeys.get("id").asText());
            assertEquals("key:name:", availableKeys.get("name").asText());

            byte[] rowKey = "row:0000000004".getBytes(StandardCharsets.UTF_8);
            JsonNode row = JSON.readTree(opened.database().get(opened.handle(), rowKey));
            assertEquals("Ada", row.get("name").asText());
            assertEquals("true", row.get("active").asText());
            assertNull(row.get("id"));
            assertEquals(List.of("row:0000000004", "row:0000000005"), JSON.convertValue(
                    JSON.readTree(opened.database().get(opened.handle(),
                            "key:id:42.0".getBytes(StandardCharsets.UTF_8))), List.class));
            assertEquals(List.of("row:0000000004"), JSON.convertValue(
                    JSON.readTree(opened.database().get(opened.handle(),
                            "key:name:Ada".getBytes(StandardCharsets.UTF_8))), List.class));
            JsonNode secondRow = JSON.readTree(opened.database().get(
                    opened.handle(), "row:0000000005".getBytes(StandardCharsets.UTF_8)));
            assertEquals("Grace", secondRow.get("name").asText());
            assertEquals(List.of("row:0000000005"), JSON.convertValue(
                    JSON.readTree(opened.database().get(opened.handle(),
                            "key:name:Grace".getBytes(StandardCharsets.UTF_8))), List.class));
        }
    }

    @Test
    void verboseProgressIsLimitedToAboutTenUpdatesForLargeSheets() {
        assertEquals(10_000, WorkbookImporter.progressInterval(50_000));
        assertEquals(25_000, WorkbookImporter.progressInterval(250_000));
        assertEquals(100_000, WorkbookImporter.progressInterval(1_000_000));
    }

    @Test
    void normalizesKeyColumnPrefixes() {
        assertEquals("customer-id", RocksDbStore.normalizeKeyColumn("  Cüstomer ID  "));
    }

    @Test
    void readsWorkbookFromStdinWithBothEngines() throws Exception {
        Path workbook = temporaryDirectory.resolve("stdin-source.xlsx");
        writeWorkbookWithGeneratedColumns(workbook, false);
        byte[] workbookBytes = Files.readAllBytes(workbook);

        for (String engine : List.of("poi", "fastexcel")) {
            Path database = temporaryDirectory.resolve("stdin-" + engine + "-db");
            StringWriter output = new StringWriter();
            CommandLine command = new CommandLine(
                    new ImportCommand(new ByteArrayInputStream(workbookBytes)));
            command.setOut(new PrintWriter(output, true));

            assertEquals(0, command.execute(
                    "-", "--engine", engine, "--db", database.toString(),
                    "--header-row", "1", "--verbose"));
            assertTrue(output.toString().contains("Reading XLSX workbook from stdin."));
            assertTrue(output.toString().contains("Opening workbook: stdin"));

            try (OpenedFamily opened = openFamily(database, "Data")) {
                JsonNode metadata = JSON.readTree(
                        opened.database().get(opened.handle(), RocksDbStore.METADATA_KEY));
                assertEquals("stdin", metadata.get("source").asText());
                assertNotNull(opened.database().get(
                        opened.handle(), "row:0000000002".getBytes(StandardCharsets.UTF_8)));
            }
        }
    }

    @Test
    void forcedEmptyHeaderGeneratesColumnNamesAndReimportRemovesStaleRows() throws Exception {
        Path workbook = temporaryDirectory.resolve("generated.xlsx");
        writeWorkbookWithGeneratedColumns(workbook, true);
        Path database = temporaryDirectory.resolve("database");

        assertEquals(0, new CommandLine(new ImportCommand()).execute(
                workbook.toString(), "--db", database.toString(), "--header-row", "1"));

        try (OpenedFamily opened = openFamily(database, "Data")) {
            assertNotNull(opened.database().get(
                    opened.handle(), "row:0000000003".getBytes(StandardCharsets.UTF_8)));
        }

        writeWorkbookWithGeneratedColumns(workbook, false);
        assertEquals(0, new CommandLine(new ImportCommand()).execute(
                workbook.toString(), "--db", database.toString(), "--header-row", "1"));

        try (OpenedFamily opened = openFamily(database, "Data")) {
            JsonNode metadata = JSON.readTree(opened.database().get(opened.handle(), RocksDbStore.METADATA_KEY));
            assertEquals(List.of("col1", "col2"), JSON.convertValue(metadata.get("columns"), List.class));
            JsonNode availableKeys = JSON.readTree(
                    opened.database().get(opened.handle(), RocksDbStore.KEYS_KEY));
            assertTrue(availableKeys.isObject());
            assertTrue(availableKeys.isEmpty());
            assertNotNull(opened.database().get(
                    opened.handle(), "row:0000000002".getBytes(StandardCharsets.UTF_8)));
            assertNull(opened.database().get(
                    opened.handle(), "row:0000000003".getBytes(StandardCharsets.UTF_8)));
        }
    }

    @Test
    void versionedImportPublishesAtomicallyAndRetainsConfiguredGenerations() throws Exception {
        Path workbook = temporaryDirectory.resolve("versioned.xlsx");
        Path database = temporaryDirectory.resolve("versioned-db");
        writeWorkbookWithGeneratedColumns(workbook, false);
        List<String> generations = new ArrayList<>();
        StringWriter output = new StringWriter();

        for (int importNumber = 0; importNumber < 3; importNumber++) {
            CommandLine command = new CommandLine(new ImportCommand());
            command.setOut(new PrintWriter(output, true));
            assertEquals(0, command.execute(workbook.toString(), "--db", database.toString(),
                    "--header-row", "1", "--key", "1", "--versioned",
                    "--keep-generations", "2", "--verbose"));
            try (OpenedFamily opened = openFamily(database, "default")) {
                JsonNode active = JSON.readTree(opened.database().get(
                        opened.handle(), RocksDbStore.ACTIVE_GENERATIONS_KEY));
                generations.add(active.get("Data").asText());
            }
        }

        assertTrue(output.toString().contains("Generation "));
        assertTrue(output.toString().contains(": published for 1 sheets."));
        assertTrue(output.toString().contains("Generation cleanup for sheet 'Data': retained 2"));
        try (OpenedFamily opened = openFamily(database, "default")) {
            JsonNode history = JSON.readTree(opened.database().get(
                    opened.handle(), RocksDbStore.GENERATION_HISTORY_KEY));
            assertEquals(List.of(generations.get(2), generations.get(1)),
                    JSON.convertValue(history.get("Data"), List.class));
        }
        try (OpenedFamily opened = openFamily(database, "Data")) {
            assertNull(opened.database().get(opened.handle(), ("generation:" + generations.get(0)
                    + ":row:0000000002").getBytes(StandardCharsets.UTF_8)));
            assertNotNull(opened.database().get(opened.handle(), ("generation:" + generations.get(1)
                    + ":row:0000000002").getBytes(StandardCharsets.UTF_8)));
            byte[] currentMetadata = ("generation:" + generations.get(2)
                    + ":__xlsx2rocksdb:metadata").getBytes(StandardCharsets.UTF_8);
            JsonNode metadata = JSON.readTree(opened.database().get(opened.handle(), currentMetadata));
            assertEquals(generations.get(2), metadata.get("generation").asText());
            String prefix = "generation:" + generations.get(2) + ":";
            JsonNode availableKeys = JSON.readTree(opened.database().get(opened.handle(),
                    (prefix + "__xlsx2rocksdb:keys").getBytes(StandardCharsets.UTF_8)));
            assertEquals(prefix + "key:col1:", availableKeys.get("col1").asText());
            assertEquals(List.of(prefix + "row:0000000002"), JSON.convertValue(
                    JSON.readTree(opened.database().get(opened.handle(),
                            (prefix + "key:col1:a").getBytes(StandardCharsets.UTF_8))), List.class));
        }
    }

    private static void writeWorkbookWithGeneratedColumns(Path path, boolean secondRow) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Data");
            sheet.createRow(0);
            var first = sheet.createRow(1);
            first.createCell(0).setCellValue("a");
            first.createCell(1).setCellValue("b");
            if (secondRow) {
                sheet.createRow(2).createCell(0).setCellValue("stale");
            }
            write(workbook, path);
        }
    }

    private static void write(XSSFWorkbook workbook, Path path) throws IOException {
        try (OutputStream output = Files.newOutputStream(path)) {
            workbook.write(output);
        }
    }

    private static List<String> familyNames(Path path) throws RocksDBException {
        try (org.rocksdb.Options options = new org.rocksdb.Options()) {
            return RocksDB.listColumnFamilies(options, path.toString()).stream()
                    .map(bytes -> new String(bytes, StandardCharsets.UTF_8))
                    .toList();
        }
    }

    private static OpenedFamily openFamily(Path path, String requested) throws RocksDBException {
        List<byte[]> names;
        try (org.rocksdb.Options options = new org.rocksdb.Options()) {
            names = RocksDB.listColumnFamilies(options, path.toString());
        }
        List<ColumnFamilyOptions> options = new ArrayList<>();
        List<ColumnFamilyDescriptor> descriptors = new ArrayList<>();
        for (byte[] name : names) {
            ColumnFamilyOptions familyOptions = new ColumnFamilyOptions();
            options.add(familyOptions);
            descriptors.add(new ColumnFamilyDescriptor(name, familyOptions));
        }
        List<ColumnFamilyHandle> handles = new ArrayList<>();
        RocksDB database;
        try (DBOptions dbOptions = new DBOptions()) {
            database = RocksDB.open(dbOptions, path.toString(), descriptors, handles);
        } finally {
            options.forEach(ColumnFamilyOptions::close);
        }
        for (int index = 0; index < names.size(); index++) {
            if (requested.equals(new String(names.get(index), StandardCharsets.UTF_8))) {
                return new OpenedFamily(database, handles.get(index), handles);
            }
        }
        handles.forEach(ColumnFamilyHandle::close);
        database.close();
        throw new IllegalArgumentException("Missing family " + requested);
    }

    private record OpenedFamily(
            RocksDB database, ColumnFamilyHandle handle, List<ColumnFamilyHandle> allHandles
    ) implements AutoCloseable {
        @Override
        public void close() {
            allHandles.forEach(ColumnFamilyHandle::close);
            database.close();
        }
    }
}
