package com.netea.xlsx2rocksdb;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.rocksdb.ColumnFamilyDescriptor;
import org.rocksdb.ColumnFamilyHandle;
import org.rocksdb.ColumnFamilyOptions;
import org.rocksdb.DBOptions;
import org.rocksdb.Options;
import org.rocksdb.RocksDB;
import org.rocksdb.RocksDBException;
import org.rocksdb.RocksIterator;
import org.rocksdb.WriteBatch;
import org.rocksdb.WriteOptions;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Locale;
import java.util.Set;
import java.util.UUID;
import java.util.function.Consumer;

final class RocksDbStore implements AutoCloseable {
    static final byte[] METADATA_KEY = "__xlsx2rocksdb:metadata".getBytes(StandardCharsets.UTF_8);
    static final byte[] KEYS_KEY = "__xlsx2rocksdb:keys".getBytes(StandardCharsets.UTF_8);
    static final byte[] ACTIVE_GENERATIONS_KEY =
            "__xlsx2rocksdb:active-generations".getBytes(StandardCharsets.UTF_8);
    static final byte[] GENERATION_HISTORY_KEY =
            "__xlsx2rocksdb:generation-history".getBytes(StandardCharsets.UTF_8);
    private static final byte[] DEFAULT_FAMILY = RocksDB.DEFAULT_COLUMN_FAMILY;
    private static final ObjectMapper JSON = new ObjectMapper();

    static {
        RocksDB.loadLibrary();
    }

    private final RocksDB database;
    private final WriteOptions writeOptions;
    private final int batchSize;
    private final Map<String, ColumnFamilyHandle> handles;
    private final String generationId;
    private final int keepGenerations;
    private final Consumer<String> progress;

    private RocksDbStore(
            RocksDB database,
            WriteOptions writeOptions,
            int batchSize,
            Map<String, ColumnFamilyHandle> handles,
            String generationId,
            int keepGenerations,
            Consumer<String> progress
    ) {
        this.database = database;
        this.writeOptions = writeOptions;
        this.batchSize = batchSize;
        this.handles = handles;
        this.generationId = generationId;
        this.keepGenerations = keepGenerations;
        this.progress = progress;
    }

    static RocksDbStore open(
            Path path,
            int batchSize,
            boolean sync,
            boolean versioned,
            int keepGenerations,
            Consumer<String> progress
    ) throws RocksDBException {
        List<byte[]> familyNames = existingFamilies(path);
        List<ColumnFamilyOptions> familyOptions = new ArrayList<>();
        List<ColumnFamilyDescriptor> descriptors = new ArrayList<>();
        List<ColumnFamilyHandle> openedHandles = new ArrayList<>();

        try {
            for (byte[] familyName : familyNames) {
                ColumnFamilyOptions options = new ColumnFamilyOptions();
                familyOptions.add(options);
                descriptors.add(new ColumnFamilyDescriptor(familyName, options));
            }
            try (DBOptions options = new DBOptions()
                    .setCreateIfMissing(true)
                    .setCreateMissingColumnFamilies(true)) {
                RocksDB database = RocksDB.open(options, path.toString(), descriptors, openedHandles);
                Map<String, ColumnFamilyHandle> handles = new LinkedHashMap<>();
                for (int index = 0; index < familyNames.size(); index++) {
                    handles.put(new String(familyNames.get(index), StandardCharsets.UTF_8), openedHandles.get(index));
                }
                String generationId = versioned ? UUID.randomUUID().toString() : null;
                RocksDbStore store = new RocksDbStore(database, new WriteOptions().setSync(sync),
                        batchSize, handles, generationId, keepGenerations, progress);
                if (versioned) {
                    progress.accept("Generation " + generationId + ": created.");
                }
                return store;
            }
        } catch (RocksDBException | RuntimeException exception) {
            openedHandles.forEach(ColumnFamilyHandle::close);
            throw exception;
        } finally {
            familyOptions.forEach(ColumnFamilyOptions::close);
        }
    }

    SheetWriter startSheet(
            String sheetName,
            String source,
            Integer headerRow,
            List<String> columns,
            List<String> keyColumns
    ) throws RocksDBException {
        ColumnFamilyHandle handle = getOrCreateFamily(sheetName);
        if (generationId == null) {
            deleteAll(handle);
        } else {
            progress.accept("Generation " + generationId + ": writing sheet '" + sheetName + "'.");
        }
        return new SheetWriter(handle, sheetName, source, headerRow,
                List.copyOf(columns), normalizedKeyColumns(keyColumns));
    }

    void publishGeneration(List<String> sheetNames) throws RocksDBException {
        if (generationId == null || sheetNames.isEmpty()) {
            return;
        }
        ColumnFamilyHandle defaultHandle = handles.get("default");
        Map<String, String> active = readStringMap(database.get(defaultHandle, ACTIVE_GENERATIONS_KEY));
        Map<String, List<String>> history = readHistory(database.get(defaultHandle, GENERATION_HISTORY_KEY));
        for (String sheetName : sheetNames) {
            active.put(sheetName, generationId);
            List<String> generations = new ArrayList<>(history.getOrDefault(sheetName, List.of()));
            generations.remove(generationId);
            generations.add(0, generationId);
            history.put(sheetName, new ArrayList<>(generations.subList(0,
                    Math.min(keepGenerations, generations.size()))));
        }
        try (WriteBatch publication = new WriteBatch()) {
            publication.put(defaultHandle, ACTIVE_GENERATIONS_KEY, toJson(active));
            publication.put(defaultHandle, GENERATION_HISTORY_KEY, toJson(history));
            database.write(writeOptions, publication);
        }
        progress.accept("Generation " + generationId + ": published for " + sheetNames.size() + " sheets.");
        for (String sheetName : sheetNames) {
            cleanupGenerations(sheetName, Set.copyOf(history.get(sheetName)));
        }
    }

    private void cleanupGenerations(String sheetName, Set<String> retained) throws RocksDBException {
        ColumnFamilyHandle handle = handles.get(sheetName);
        int removed = 0;
        try (RocksIterator iterator = database.newIterator(handle); WriteBatch deletes = new WriteBatch()) {
            for (iterator.seekToFirst(); iterator.isValid(); iterator.next()) {
                String key = new String(iterator.key(), StandardCharsets.UTF_8);
                String keyGeneration = generationFromKey(key);
                boolean reservedManifest = handle == handles.get("default")
                        && (key.equals(new String(ACTIVE_GENERATIONS_KEY, StandardCharsets.UTF_8))
                        || key.equals(new String(GENERATION_HISTORY_KEY, StandardCharsets.UTF_8)));
                if (!reservedManifest && (keyGeneration == null || !retained.contains(keyGeneration))) {
                    deletes.delete(handle, iterator.key().clone());
                    removed++;
                    if (removed % batchSize == 0) {
                        database.write(writeOptions, deletes);
                        deletes.clear();
                    }
                }
            }
            if (removed % batchSize != 0) {
                database.write(writeOptions, deletes);
            }
            iterator.status();
        }
        progress.accept("Generation cleanup for sheet '" + sheetName + "': retained "
                + retained.size() + ", removed " + removed + " entries.");
    }

    private static String generationFromKey(String key) {
        if (!key.startsWith("generation:")) {
            return null;
        }
        int end = key.indexOf(':', "generation:".length());
        return end < 0 ? null : key.substring("generation:".length(), end);
    }

    @SuppressWarnings("unchecked")
    private static Map<String, String> readStringMap(byte[] value) {
        if (value == null) {
            return new LinkedHashMap<>();
        }
        try {
            return JSON.readValue(value, LinkedHashMap.class);
        } catch (IOException exception) {
            throw new IllegalArgumentException("Could not read active-generation manifest", exception);
        }
    }

    private static Map<String, List<String>> readHistory(byte[] value) {
        Map<String, List<String>> result = new LinkedHashMap<>();
        if (value == null) {
            return result;
        }
        try {
            JsonNode root = JSON.readTree(value);
            root.fields().forEachRemaining(entry -> {
                List<String> generations = new ArrayList<>();
                entry.getValue().forEach(item -> generations.add(item.asText()));
                result.put(entry.getKey(), generations);
            });
            return result;
        } catch (IOException exception) {
            throw new IllegalArgumentException("Could not read generation history", exception);
        }
    }

    private ColumnFamilyHandle getOrCreateFamily(String name) throws RocksDBException {
        ColumnFamilyHandle existing = handles.get(name);
        if (existing != null) {
            return existing;
        }
        try (ColumnFamilyOptions options = new ColumnFamilyOptions()) {
            ColumnFamilyHandle created = database.createColumnFamily(
                    new ColumnFamilyDescriptor(name.getBytes(StandardCharsets.UTF_8), options));
            handles.put(name, created);
            return created;
        }
    }

    private void deleteAll(ColumnFamilyHandle handle) throws RocksDBException {
        try (RocksIterator iterator = database.newIterator(handle); WriteBatch batch = new WriteBatch()) {
            int operations = 0;
            for (iterator.seekToFirst(); iterator.isValid(); iterator.next()) {
                batch.delete(handle, iterator.key().clone());
                operations++;
                if (operations >= batchSize) {
                    database.write(writeOptions, batch);
                    batch.clear();
                    operations = 0;
                }
            }
            if (operations > 0) {
                database.write(writeOptions, batch);
            }
            iterator.status();
        }
    }

    private static List<byte[]> existingFamilies(Path path) throws RocksDBException {
        if (!Files.exists(path.resolve("CURRENT"))) {
            return List.of(DEFAULT_FAMILY);
        }
        try (Options options = new Options()) {
            return RocksDB.listColumnFamilies(options, path.toString());
        }
    }

    private byte[] rowKey(int sourceRow) {
        return ("row:%010d".formatted(sourceRow)).getBytes(StandardCharsets.UTF_8);
    }

    static String normalizeKeyColumn(String column) {
        String normalized = Normalizer.normalize(column, Normalizer.Form.NFKD)
                .replaceAll("\\p{M}+", "")
                .toLowerCase(Locale.ROOT)
                .replaceAll("[^a-z0-9]+", "-")
                .replaceAll("^-|-$", "");
        if (normalized.isEmpty()) {
            throw new IllegalArgumentException("Key column name cannot be normalized: '" + column + "'");
        }
        return normalized;
    }

    private static Map<String, String> normalizedKeyColumns(List<String> keyColumns) {
        Map<String, String> normalized = new LinkedHashMap<>();
        for (String column : keyColumns) {
            String prefix = normalizeKeyColumn(column);
            String existing = normalized.putIfAbsent(column, prefix);
            if (existing != null) {
                continue;
            }
            if (normalized.entrySet().stream().anyMatch(entry ->
                    !entry.getKey().equals(column) && entry.getValue().equals(prefix))) {
                throw new IllegalArgumentException(
                        "Key columns normalize to the same prefix: " + keyColumns);
            }
        }
        return Map.copyOf(normalized);
    }

    private byte[] indexKey(String normalizedColumn, String value) {
        return ("key:" + normalizedColumn + ":" + value).getBytes(StandardCharsets.UTF_8);
    }

    private byte[] scopedKey(byte[] key) {
        if (generationId == null) {
            return key;
        }
        return ("generation:" + generationId + ":" + new String(key, StandardCharsets.UTF_8))
                .getBytes(StandardCharsets.UTF_8);
    }

    private static byte[] toJson(Object value) {
        try {
            return JSON.writeValueAsBytes(value);
        } catch (JsonProcessingException exception) {
            throw new IllegalArgumentException("Could not serialize row as JSON", exception);
        }
    }

    final class SheetWriter implements AutoCloseable {
        private final ColumnFamilyHandle handle;
        private final String sheetName;
        private final String source;
        private final Integer headerRow;
        private final List<String> columns;
        private final Map<String, String> keyColumns;
        private final Instant importedAt = Instant.now();
        private final WriteBatch batch = new WriteBatch();
        private final Map<String, PendingIndex> pendingIndexes = new LinkedHashMap<>();
        private int pendingOperations;
        private int rowCount;
        private boolean finished;
        private boolean closed;

        private SheetWriter(
                ColumnFamilyHandle handle,
                String sheetName,
                String source,
                Integer headerRow,
                List<String> columns,
                Map<String, String> keyColumns
        ) {
            this.handle = handle;
            this.sheetName = sheetName;
            this.source = source;
            this.headerRow = headerRow;
            this.columns = columns;
            this.keyColumns = keyColumns;
        }

        void putRow(
                int sourceRow, Map<String, String> values, Map<String, String> keyValues
        ) throws RocksDBException {
            ensureWritable();
            byte[] rowKey = scopedKey(rowKey(sourceRow));
            String rowKeyText = new String(rowKey, StandardCharsets.UTF_8);
            for (Map.Entry<String, String> keyColumn : keyColumns.entrySet()) {
                String value = keyValues.getOrDefault(keyColumn.getKey(), "");
                if (value.isBlank()) {
                    continue;
                }
                byte[] indexKey = scopedKey(indexKey(keyColumn.getValue(), value));
                String indexKeyText = new String(indexKey, StandardCharsets.UTF_8);
                pendingIndexes
                        .computeIfAbsent(indexKeyText, ignored -> new PendingIndex(indexKey, new ArrayList<>()))
                        .rowKeys()
                        .add(rowKeyText);
                pendingOperations++;
            }

            batch.put(handle, rowKey, toJson(values));
            pendingOperations++;
            rowCount++;
            if (pendingOperations >= batchSize) {
                flush();
            }
        }

        void finish() throws RocksDBException {
            ensureWritable();
            Map<String, Object> metadata = new LinkedHashMap<>();
            metadata.put("sheet", sheetName);
            metadata.put("source", source);
            metadata.put("importedAt", importedAt.toString());
            metadata.put("headerRow", headerRow);
            metadata.put("columns", columns);
            metadata.put("keyColumns", keyColumns);
            metadata.put("rowCount", rowCount);
            if (generationId != null) {
                metadata.put("generation", generationId);
            }
            batch.put(handle, scopedKey(METADATA_KEY), toJson(metadata));
            pendingOperations++;
            Map<String, String> availableKeys = new LinkedHashMap<>();
            keyColumns.forEach((column, normalized) ->
                    availableKeys.put(column, keyPrefix(normalized)));
            batch.put(handle, scopedKey(KEYS_KEY), toJson(availableKeys));
            pendingOperations++;
            flush();
            finished = true;
        }

        private String keyPrefix(String normalized) {
            return (generationId == null ? "" : "generation:" + generationId + ":")
                    + "key:" + normalized + ":";
        }

        private void flush() throws RocksDBException {
            if (pendingOperations == 0) {
                return;
            }
            for (PendingIndex index : pendingIndexes.values()) {
                List<String> rowKeys = readIndexRows(database.get(handle, index.key()));
                rowKeys.addAll(index.rowKeys());
                batch.put(handle, index.key(), toJson(rowKeys));
            }
            database.write(writeOptions, batch);
            batch.clear();
            pendingIndexes.clear();
            pendingOperations = 0;
        }

        private List<String> readIndexRows(byte[] storedValue) {
            List<String> rowKeys = new ArrayList<>();
            if (storedValue == null || storedValue.length == 0) {
                return rowKeys;
            }
            try {
                JsonNode value = JSON.readTree(storedValue);
                if (value.isArray()) {
                    value.forEach(item -> rowKeys.add(item.asText()));
                    return rowKeys;
                }
                if (value.isTextual()) {
                    rowKeys.add(value.asText());
                    return rowKeys;
                }
            } catch (IOException ignored) {
                // Support index values written by versions that stored a plain row key.
            }
            rowKeys.add(new String(storedValue, StandardCharsets.UTF_8));
            return rowKeys;
        }

        private void ensureWritable() {
            if (closed || finished) {
                throw new IllegalStateException("Sheet writer is already closed");
            }
        }

        @Override
        public void close() {
            if (!closed) {
                batch.close();
                closed = true;
            }
        }
    }

    private record PendingIndex(byte[] key, List<String> rowKeys) {
    }

    @Override
    public void close() {
        handles.values().forEach(ColumnFamilyHandle::close);
        writeOptions.close();
        database.close();
    }
}
