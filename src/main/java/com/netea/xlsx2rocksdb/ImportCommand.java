package com.netea.xlsx2rocksdb;

import org.rocksdb.RocksDBException;
import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;
import picocli.CommandLine.Parameters;
import picocli.CommandLine.Spec;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.concurrent.Callable;
import java.util.function.Consumer;

@Command(
        name = "xlsx2rocksdb",
        mixinStandardHelpOptions = true,
        version = "xlsx2rocksdb 0.1.0",
        description = "Imports XLSX worksheets into RocksDB column families."
)
public final class ImportCommand implements Callable<Integer> {
    private final InputStream standardInput;

    public ImportCommand() {
        this(System.in);
    }

    ImportCommand(InputStream standardInput) {
        this.standardInput = standardInput;
    }

    @Spec
    private CommandLine.Model.CommandSpec spec;

    @Parameters(index = "0", paramLabel = "FILE.xlsx|-",
            description = "XLSX workbook to import, or '-' to read it from stdin.")
    private Path input;

    @Option(names = {"-d", "--db"}, paramLabel = "DIRECTORY",
            description = "RocksDB directory (default: FILE.rocksdb beside the workbook).")
    private Path database;

    @Option(names = {"-s", "--sheets"}, split = ",", paramLabel = "NAME",
            description = "Sheet names to process; repeat or comma-separate (default: all sheets).")
    private List<String> sheets = new ArrayList<>();

    @Option(names = {"-H", "--header-row"}, paramLabel = "ROW", arity = "1",
            description = "One-based header row for every selected sheet (default: first non-empty row).")
    private Integer headerRow;

    @Option(names = {"-c", "--columns"}, split = ",", paramLabel = "SELECTOR",
            description = "Columns to import by name, index:N, or letter:A; repeat or comma-separate.")
    private List<String> columns = new ArrayList<>();

    @Option(names = {"-k", "--key"}, split = ",", paramLabel = "SELECTOR",
            description = "Key/index columns by name, one-based number, index:N, or letter:A; repeat or comma-separate.")
    private List<String> keyColumns = new ArrayList<>();

    @Option(names = "--engine", paramLabel = "ENGINE", defaultValue = "poi",
            converter = ExcelEngineConverter.class,
            description = "Excel reader: poi or fastexcel (default: ${DEFAULT-VALUE}).")
    private ExcelEngine engine;

    @Option(names = "--batch-size", paramLabel = "ROWS", defaultValue = "1000",
            description = "Maximum RocksDB operations per write batch (default: ${DEFAULT-VALUE}).")
    private int batchSize;

    @Option(names = "--sync", description = "Synchronously flush each RocksDB write batch to disk.")
    private boolean sync;

    @Option(names = "--versioned",
            description = "Import into a new generation and publish it atomically after completion.")
    private boolean versioned;

    @Option(names = "--keep-generations", paramLabel = "N", defaultValue = "2",
            description = "Generations retained per processed sheet in versioned mode (default: ${DEFAULT-VALUE}).")
    private int keepGenerations;

    @Option(names = {"-v", "--verbose"}, description = "Report concise import steps and progress to stdout.")
    private boolean verbose;

    @Override
    public Integer call() {
        validateArguments();
        Path dbPath = database != null ? database.toAbsolutePath().normalize() : defaultDatabasePath(input);
        Set<String> selectedSheets = normalizedSheets();

        try (StagedWorkbook staged = stageWorkbook()) {
            report("Opening RocksDB: " + dbPath);
            try (RocksDbStore store = RocksDbStore.open(
                    dbPath, batchSize, sync, versioned, keepGenerations, progressReporter())) {
                report("Opening workbook: " + staged.source());
                report("Excel engine: " + engine.name().toLowerCase());
                List<WorkbookImporter.SheetResult> results = engine == ExcelEngine.POI
                        ? new WorkbookImporter().importWorkbook(
                                staged.path(), staged.source(), selectedSheets, headerRow, columns, keyColumns,
                                store, progressReporter(), warningReporter())
                        : new FastExcelWorkbookImporter().importWorkbook(
                                staged.path(), staged.source(), selectedSheets, headerRow, columns, keyColumns,
                                store, progressReporter(), warningReporter());
                store.publishGeneration(results.stream().map(WorkbookImporter.SheetResult::sheetName).toList());
                report("Workbook import complete.");

                int totalRows = results.stream().mapToInt(WorkbookImporter.SheetResult::rowCount).sum();
                for (WorkbookImporter.SheetResult result : results) {
                    spec.commandLine().getOut().printf(
                            "Imported sheet '%s': %d rows, %d columns, header row %s%n",
                            result.sheetName(), result.rowCount(), result.columnCount(),
                            result.headerRow() == null ? "none" : result.headerRow());
                }
                spec.commandLine().getOut().printf(
                        "Finished: %d sheets, %d rows -> %s%n", results.size(), totalRows, dbPath);
                return CommandLine.ExitCode.OK;
            }
        } catch (IOException | RocksDBException | IllegalArgumentException exception) {
            throw new CommandLine.ParameterException(spec.commandLine(), exception.getMessage(), exception);
        }
    }

    private StagedWorkbook stageWorkbook() throws IOException {
        if (!isStandardInput()) {
            Path path = input.toAbsolutePath().normalize();
            return new StagedWorkbook(path, path.toString(), false);
        }

        report("Reading XLSX workbook from stdin.");
        Path temporary = Files.createTempFile("xlsx2rocksdb-stdin-", ".xlsx");
        try {
            Files.copy(standardInput, temporary, StandardCopyOption.REPLACE_EXISTING);
            return new StagedWorkbook(temporary, "stdin", true);
        } catch (IOException | RuntimeException exception) {
            Files.deleteIfExists(temporary);
            throw exception;
        }
    }

    private Consumer<String> progressReporter() {
        return this::report;
    }

    private Consumer<String> warningReporter() {
        return message -> spec.commandLine().getErr().println("Warning: " + message);
    }

    private void report(String message) {
        if (verbose) {
            spec.commandLine().getOut().println(message);
        }
    }

    private void validateArguments() {
        if (input == null || (!isStandardInput() && !Files.isRegularFile(input))) {
            throw new CommandLine.ParameterException(spec.commandLine(), "Input file does not exist: " + input);
        }
        String fileName = input.getFileName().toString().toLowerCase();
        if (!isStandardInput() && !fileName.endsWith(".xlsx")) {
            throw new CommandLine.ParameterException(spec.commandLine(), "Input must have an .xlsx extension: " + input);
        }
        if (headerRow != null && headerRow < 1) {
            throw new CommandLine.ParameterException(spec.commandLine(), "--header-row must be at least 1");
        }
        if (batchSize < 1) {
            throw new CommandLine.ParameterException(spec.commandLine(), "--batch-size must be at least 1");
        }
        if (keepGenerations < 1) {
            throw new CommandLine.ParameterException(spec.commandLine(), "--keep-generations must be at least 1");
        }
    }

    private boolean isStandardInput() {
        return input != null && "-".equals(input.toString());
    }

    private Set<String> normalizedSheets() {
        Set<String> selected = new LinkedHashSet<>();
        for (String sheet : sheets) {
            String name = sheet.trim();
            if (name.isEmpty()) {
                throw new CommandLine.ParameterException(spec.commandLine(), "Sheet names must not be empty");
            }
            selected.add(name);
        }
        return selected;
    }

    private static Path defaultDatabasePath(Path workbook) {
        if ("-".equals(workbook.toString())) {
            return Path.of("stdin.rocksdb").toAbsolutePath().normalize();
        }
        Path absolute = workbook.toAbsolutePath().normalize();
        String fileName = absolute.getFileName().toString();
        String baseName = fileName.substring(0, fileName.length() - ".xlsx".length());
        return absolute.resolveSibling(baseName + ".rocksdb");
    }

    private record StagedWorkbook(Path path, String source, boolean temporary) implements AutoCloseable {
        @Override
        public void close() throws IOException {
            if (temporary) {
                Files.deleteIfExists(path);
            }
        }
    }

    enum ExcelEngine {
        POI, FASTEXCEL
    }

    static final class ExcelEngineConverter implements CommandLine.ITypeConverter<ExcelEngine> {
        @Override
        public ExcelEngine convert(String value) {
            try {
                return ExcelEngine.valueOf(value.toUpperCase());
            } catch (IllegalArgumentException exception) {
                throw new CommandLine.TypeConversionException(
                        "expected 'poi' or 'fastexcel' but was '" + value + "'");
            }
        }
    }
}
