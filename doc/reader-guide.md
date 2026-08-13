# Reading xlsx2rocksdb databases

This document defines the storage contract produced by `xlsx2rocksdb` and the recommended Java reader design. It covers ordinary and versioned imports, worksheet column families, secondary indexes, snapshots, and concurrent access while the importer is running.

## Process and concurrency model

RocksDB is embedded rather than a database server. Only one process may hold the primary read/write instance for a database directory. During an import, `xlsx2rocksdb` is that primary process.

A separate, long-running server must therefore open the database with `RocksDB.openAsSecondary(...)`, not `RocksDB.open(...)` or `openReadOnly(...)`. One secondary `RocksDB` object may be shared by the server's reader threads; RocksDB point reads and iterators are thread-safe. Give every server instance its own writable secondary directory for its logs and local state.

The server controls when new primary data becomes visible by calling `tryCatchUpWithPrimary()`. A practical policy is to call it on a timer or before beginning a request that requires fresh data. Catch-up is best effort, so production code should report failures and retry. A secondary opened before the primary creates a new column family does not discover that family dynamically; reopen the secondary when new worksheet names are introduced.

Do not use a normal read-only instance concurrently with the importer. Read-only mode supports multiple processes only while no primary is writing.

## Column families and JSON encoding

Each worksheet is stored in a column family whose name exactly matches the Excel sheet name. All application keys and JSON values use UTF-8.

Rows are JSON objects. Column values are strings because the importers store the formatted or textual representation read from Excel:

```json
{"customer_id":"42","name":"Ada","active":"true"}
```

The following logical keys exist inside each worksheet column family:

| Key | JSON value |
| --- | --- |
| `row:0000000004` | Row object; the number is the one-based source Excel row. |
| `key:<normalized-column>:<cell-value>` | Array of matching row-key strings, including a one-element array for a unique value. |
| `__xlsx2rocksdb:metadata` | Import source, timestamp, header row, selected columns, key columns, and row count. |
| `__xlsx2rocksdb:keys` | Map from original column names to their complete index prefixes. |

Example index discovery and lookup values:

```json
{"Customer ID":"key:customer-id:","Country":"key:country:"}
```

```text
key:country:DE
```

```json
["row:0000000004","row:0000000021"]
```

Column-name normalization applies only to the index prefix: accents are removed, letters are lowercased, and punctuation or whitespace becomes `-`. The cell value after the final prefix separator is preserved exactly as the Excel engine read it. Readers should obtain the prefix from `__xlsx2rocksdb:keys` instead of reimplementing normalization. Blank key cells have no index entry.

## Versioned layout

With `--versioned`, every worksheet key described above receives this prefix:

```text
generation:<uuid>:
```

For example:

```text
generation:a81f...:__xlsx2rocksdb:metadata
generation:a81f...:__xlsx2rocksdb:keys
generation:a81f...:row:0000000004
generation:a81f...:key:customer-id:42
```

Index arrays contain complete generation-prefixed row keys:

```json
["generation:a81f...:row:0000000004"]
```

The default column family contains two unprefixed manifests:

```text
__xlsx2rocksdb:active-generations
__xlsx2rocksdb:generation-history
```

The active manifest maps each versioned worksheet to the generation readers must use:

```json
{
  "Customers":"a81f...",
  "Orders":"bf20..."
}
```

The history manifest maps sheets to generation IDs in newest-first order:

```json
{
  "Customers":["a81f...","previous..."],
  "Orders":["bf20...","previous..."]
}
```

Publication replaces both manifests in one atomic RocksDB write batch after every selected sheet has been imported successfully. Use only `active-generations` for normal routing; `generation-history` is diagnostic and retention metadata. Unselected sheets keep their previous active generation.

If a sheet is absent from the active manifest, treat it as unversioned and use an empty prefix. This permits a database to contain versioned and legacy sheets during migration.

## Consistent request algorithm

A request that performs more than one RocksDB read must use one snapshot. Without it, one `get()` could resolve an old generation while a later `get()` observes the newly published manifest.

For each request or logical read transaction:

1. Optionally call `tryCatchUpWithPrimary()` before creating the snapshot.
2. Acquire a snapshot from the secondary instance.
3. Read `__xlsx2rocksdb:active-generations` from the default column family through the snapshot.
4. Resolve the generation prefix for the requested sheet.
5. Read metadata, key discovery, index arrays, and rows through the same snapshot.
6. Release the snapshot promptly.

Never call catch-up halfway through a request and replace its snapshot. Different requests may legitimately observe different published generations.

### Lookup by secondary key

To look up `Customer ID = 42` in `Customers`:

1. Resolve the active prefix, such as `generation:a81f...:`.
2. Read `<prefix>__xlsx2rocksdb:keys` from the `Customers` column family.
3. Read the prefix registered under the exact column name `Customer ID`.
4. Append the exact lookup value `42` to form the RocksDB index key.
5. Decode its JSON array of row keys.
6. Fetch every row key in that array through the same snapshot.

A missing index key means no matching rows. An existing index whose array references a missing row is an integrity error and should be logged rather than silently interpreted as no match.

### Scanning all rows

Create an iterator with the request's `ReadOptions`, seek to `<prefix>row:`, and advance while keys still start with that byte prefix. Do not iterate the complete column family and assume every value is a row; it also contains metadata, indexes, retained generations, and possibly manifests when the worksheet is named `default`.

## RocksJava reader skeleton

Use the same RocksJava version as the importer (`org.rocksdb:rocksdbjni:10.10.1.1`). The example omits application-specific exception mapping and JSON DTOs but shows the required lifecycle.

```java
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.rocksdb.*;

import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.*;
import java.util.function.Function;

public final class XlsxRocksReader implements AutoCloseable {
    private static final ObjectMapper JSON = new ObjectMapper();
    private static final byte[] ACTIVE = bytes("__xlsx2rocksdb:active-generations");

    static {
        RocksDB.loadLibrary();
    }

    private final RocksDB db;
    private final DBOptions dbOptions;
    private final List<ColumnFamilyOptions> familyOptions;
    private final List<ColumnFamilyHandle> allHandles;
    private final Map<String, ColumnFamilyHandle> handles;

    public XlsxRocksReader(Path primaryDirectory, Path secondaryDirectory)
            throws RocksDBException {
        List<byte[]> names;
        try (Options options = new Options()) {
            names = RocksDB.listColumnFamilies(options, primaryDirectory.toString());
        }

        this.dbOptions = new DBOptions();
        this.familyOptions = new ArrayList<>();
        List<ColumnFamilyDescriptor> descriptors = new ArrayList<>();
        for (byte[] name : names) {
            ColumnFamilyOptions options = new ColumnFamilyOptions();
            familyOptions.add(options);
            descriptors.add(new ColumnFamilyDescriptor(name, options));
        }

        this.allHandles = new ArrayList<>();
        this.db = RocksDB.openAsSecondary(
                dbOptions,
                primaryDirectory.toString(),
                secondaryDirectory.toString(),
                descriptors,
                allHandles);
        this.handles = new HashMap<>();
        for (int i = 0; i < names.size(); i++) {
            handles.put(new String(names.get(i), StandardCharsets.UTF_8), allHandles.get(i));
        }
    }

    // Serialize calls from a scheduler if several threads can initiate catch-up.
    public synchronized void catchUp() throws RocksDBException {
        db.tryCatchUpWithPrimary();
    }

    public List<Map<String, String>> findByKey(
            String sheet, String column, String value) throws RocksDBException {
        return withSnapshot(readOptions -> {
            try {
                ColumnFamilyHandle sheetHandle = requireHandle(sheet);
                String generationPrefix = generationPrefix(sheet, readOptions);

                byte[] discoveryBytes = db.get(sheetHandle, readOptions,
                        bytes(generationPrefix + "__xlsx2rocksdb:keys"));
                if (discoveryBytes == null) {
                    throw new IllegalStateException("Missing key metadata for sheet " + sheet);
                }
                Map<String, String> availableKeys = JSON.readValue(
                        discoveryBytes, new TypeReference<>() {});
                String indexPrefix = availableKeys.get(column);
                if (indexPrefix == null) {
                    throw new IllegalArgumentException(
                            "Column is not configured as a key: " + column);
                }

                byte[] indexBytes = db.get(sheetHandle, readOptions, bytes(indexPrefix + value));
                if (indexBytes == null) {
                    return List.of();
                }
                List<String> rowKeys = JSON.readValue(indexBytes, new TypeReference<>() {});
                List<Map<String, String>> rows = new ArrayList<>(rowKeys.size());
                for (String rowKey : rowKeys) {
                    byte[] rowBytes = db.get(sheetHandle, readOptions, bytes(rowKey));
                    if (rowBytes == null) {
                        throw new IllegalStateException("Index references missing row " + rowKey);
                    }
                    rows.add(JSON.readValue(rowBytes, new TypeReference<>() {}));
                }
                return List.copyOf(rows);
            } catch (RocksDBException exception) {
                throw new ReaderFailure(exception);
            } catch (java.io.IOException exception) {
                throw new IllegalStateException("Invalid xlsx2rocksdb JSON", exception);
            }
        });
    }

    private String generationPrefix(String sheet, ReadOptions readOptions)
            throws RocksDBException, java.io.IOException {
        byte[] manifestBytes = db.get(requireHandle("default"), readOptions, ACTIVE);
        if (manifestBytes == null) {
            return ""; // Entire database predates versioned imports.
        }
        Map<String, String> active = JSON.readValue(manifestBytes, new TypeReference<>() {});
        String generation = active.get(sheet);
        return generation == null ? "" : "generation:" + generation + ":";
    }

    private <T> T withSnapshot(Function<ReadOptions, T> operation) throws RocksDBException {
        Snapshot snapshot = db.getSnapshot();
        try (ReadOptions readOptions = new ReadOptions().setSnapshot(snapshot)) {
            try {
                return operation.apply(readOptions);
            } catch (ReaderFailure failure) {
                throw failure.cause;
            }
        } finally {
            db.releaseSnapshot(snapshot);
        }
    }

    private ColumnFamilyHandle requireHandle(String name) {
        ColumnFamilyHandle handle = handles.get(name);
        if (handle == null) {
            throw new IllegalArgumentException("Unknown worksheet: " + name);
        }
        return handle;
    }

    private static byte[] bytes(String value) {
        return value.getBytes(StandardCharsets.UTF_8);
    }

    private static final class ReaderFailure extends RuntimeException {
        private final RocksDBException cause;

        private ReaderFailure(RocksDBException cause) {
            super(cause);
            this.cause = cause;
        }
    }

    @Override
    public void close() {
        allHandles.forEach(ColumnFamilyHandle::close);
        db.close();
        familyOptions.forEach(ColumnFamilyOptions::close);
        dbOptions.close();
    }
}
```

Call `catchUp()` outside a request snapshot. The server may then run many `findByKey` calls concurrently on the shared secondary instance; each call receives its own snapshot and `ReadOptions`.

## Operational guidance

- Keep the primary database and every secondary state directory on storage accessible to the corresponding process. Never share one secondary directory between server processes.
- Generation retention primarily provides rollback/inspection history and tolerance for lagging consumers. Correct requests must still use snapshots; do not rely on retention alone for consistency.
- Release iterators, snapshots, `ReadOptions`, column-family handles, options, and the database deterministically.
- Monitor catch-up failures, corrupt JSON, missing manifests, missing row references, and unknown column families.
- Treat generation IDs and physical row numbers as opaque. They are routing identifiers, not stable business IDs.
- If freshness matters, expose the active generation observed by a request in diagnostics so lagging secondary instances are visible.

RocksDB references: [multi-process access](https://github.com/facebook/rocksdb/wiki/RocksDB-FAQ), [atomic batches and snapshots](https://github.com/facebook/rocksdb/wiki/Basic-Operations), and [RocksJava API](https://javadoc.io/doc/org.rocksdb/rocksdbjni/10.10.1.1/).
