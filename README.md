# xlsx2rocksdb

[![CI](https://github.com/neteagmbh/xlsx2rocksdb/actions/workflows/ci.yml/badge.svg)](https://github.com/neteagmbh/xlsx2rocksdb/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Java 17+](https://img.shields.io/badge/Java-17%2B-blue.svg)](https://adoptium.net/)

`xlsx2rocksdb` imports each selected XLSX worksheet into a RocksDB column family. A column family is used as the RocksDB equivalent of a table, and its name is exactly the Excel sheet name.

> **Project status:** `0.1.0` is a prerelease. Test it with representative workbooks and retain backups before adopting it in production.

See [Reading xlsx2rocksdb databases](doc/reader-guide.md) for the storage contract and a RocksJava secondary-reader implementation.

## Features

- Streams large XLSX workbooks without loading complete worksheets into memory.
- Supports Apache POI and fastexcel reader engines.
- Selects sheets, headers, columns, and secondary-index columns from the CLI.
- Stores rows and index values as JSON in per-sheet RocksDB column families.
- Supports atomic generation publication for consistent long-running readers.
- Reads workbooks from files or standard input.

## Requirements

- Java 17 or newer
- Maven 3.9 or newer

## Build

```bash
mvn clean package
```

This creates the executable fat JAR `target/xlsx2rocksdb.jar`.

## Installation

Download `xlsx2rocksdb.jar` and its SHA-256 checksum from the [GitHub Releases](https://github.com/neteagmbh/xlsx2rocksdb/releases) page, then verify it:

```bash
sha256sum --check xlsx2rocksdb.jar.sha256
java -jar xlsx2rocksdb.jar --version
```

On macOS, use `shasum -a 256 xlsx2rocksdb.jar` and compare the result with the published checksum.

## Usage

```bash
java -jar target/xlsx2rocksdb.jar [OPTIONS] FILE.xlsx
```

Examples:

```bash
# Import every sheet to workbook.rocksdb
java -jar target/xlsx2rocksdb.jar workbook.xlsx

# Read the XLSX stream from stdin (spooled temporarily for two-pass parsing)
cat workbook.xlsx | java -jar target/xlsx2rocksdb.jar - --db ./data/stdin-db

# Import two sheets to a specific database directory
java -jar target/xlsx2rocksdb.jar workbook.xlsx \
  --db ./data/customer-db \
  --sheets Customers,Orders

# Force row 4 to be the header and import a subset of columns
java -jar target/xlsx2rocksdb.jar workbook.xlsx \
  --header-row 4 \
  --columns id,name,letter:F,index:8

# Use fastexcel and create secondary indexes for two key columns
java -jar target/xlsx2rocksdb.jar workbook.xlsx \
  --engine fastexcel \
  --key 1,2
```

Options:

| Option | Meaning |
| --- | --- |
| `-d, --db DIRECTORY` | RocksDB directory. Defaults to `<workbook-name>.rocksdb` beside the input file, or `stdin.rocksdb` in the current directory for input `-`. |
| `-s, --sheets NAME` | Sheet names to import. Repeat the option or comma-separate names. Defaults to all sheets. |
| `-H, --header-row ROW` | One-based header row used for every selected sheet. Defaults to the first non-empty row. |
| `-c, --columns SELECTOR` | Column subset. Use a header name, `name:...`, `index:N` (one-based), or `letter:A`. Repeat or comma-separate selectors. |
| `-k, --key SELECTOR` | Secondary-index columns. A bare number is always a one-based column index; names, `name:...`, `index:N`, and `letter:A` are also supported. Repeat or comma-separate. Key columns need not be included in `--columns`. |
| `--engine ENGINE` | XLSX reader: `poi` (default) or `fastexcel`. |
| `--batch-size ROWS` | Maximum RocksDB operations per write batch. Defaults to 1000. |
| `--sync` | Ask RocksDB to synchronously flush every write batch to disk. |
| `--versioned` | Write an isolated generation and publish it atomically after every selected sheet succeeds. |
| `--keep-generations N` | Retain the newest `N` generations per processed sheet in versioned mode. Defaults to `2`. |
| `-v, --verbose` | Report concise opening, per-sheet analysis, preparation, and progress steps to stdout. |
| `-h, --help` | Show command help. |
| `-V, --version` | Show the version. |

Bare `--columns` selectors first match an exact header name. If no header matches, digits are treated as a one-based index and letters as an Excel column reference. For `--key`, a bare numeric selector is always a one-based index, while the resolved header name—not the number—is used in the RocksDB key prefix. Use `name:1` to select a header literally named `1`.

## Import behavior

- When `--header-row` is omitted, the first row containing a non-blank displayed value is the header.
- Blank header cells are named `col1`, `col2`, and so on. Duplicate names receive `_2`, `_3`, and so on.
- If a forced header row is empty, columns are named `col1`, `col2`, and so on based on the widest following data row.
- Completely empty sheets create an empty column family with metadata and no columns.
- Empty data rows are skipped. Cell values are stored as their Excel-formatted text; formulas use the cached result last saved by Excel.
- Filename `-` reads the XLSX bytes from stdin. Because both engines make two passes, the stream is copied to a temporary `.xlsx` file and deleted after import; metadata records the source as `stdin`.
- XLSX worksheets are read with POI's SAX/Event API and a read-only shared-string table. Sheets are scanned once for header/width analysis and once for import, so worksheet rows are never loaded into the heap as a complete object model. Rows are written directly to RocksDB batches; remaining memory use is primarily styles, shared strings, and the active batch.
- The `fastexcel` engine is an alternative file-backed streaming reader. It is generally faster and has a simpler text conversion model; unlike POI, it does not reproduce Excel display formatting from styles.
- Verbose progress reports approximately every 10% on large sheets, but never more often than every 10,000 source rows. It does not log individual rows.
- Re-importing a processed sheet replaces that sheet's existing keys, so rows removed from Excel do not remain in RocksDB. Column families for unselected sheets are unchanged.
- With `--versioned`, rows, indexes, and metadata are stored below `generation:<uuid>:` prefixes. The active-generation manifest is published atomically only after the complete import succeeds; verbose output reports creation, publication, and cleanup. Older and abandoned generations beyond `--keep-generations` are then removed.

## RocksDB layout

Each imported sheet column family contains:

- `__xlsx2rocksdb:metadata`: JSON containing the source path, import time, header row, selected columns, and row count.
- `__xlsx2rocksdb:keys`: JSON mapping available key-column names to their complete RocksDB lookup prefixes, for example `{"customer_id":"key:customer-id:"}`. It is `{}` when no key columns were requested.
- `row:0000000002`, `row:0000000003`, ...: JSON objects keyed by their one-based source Excel row number.
- `key:<normalized-column>:<value>`: for every non-blank `--key` cell, contains a JSON array of matching `row:...` keys, including a one-element array for a unique value. Column prefixes are lowercase ASCII with punctuation/whitespace collapsed to `-`; values are preserved exactly as read. Duplicate values append rows to the array, and all rows are imported.

The default column family remains present because RocksDB requires it. A worksheet named `default` intentionally uses that family.

### Versioned RocksDB layout

With `--versioned`, the logical keys above are scoped by an import-generation UUID:

```text
generation:<uuid>:__xlsx2rocksdb:metadata
generation:<uuid>:__xlsx2rocksdb:keys
generation:<uuid>:row:0000000002
generation:<uuid>:key:<normalized-column>:<value>
```

For example, an index entry and its value look like:

```text
generation:a81f...:key:customer-id:42
```

```json
["generation:a81f...:row:0000000004"]
```

The generation's `__xlsx2rocksdb:keys` value contains complete lookup prefixes, so readers do not need to construct or normalize column prefixes themselves:

```json
{
  "Customer ID": "generation:a81f...:key:customer-id:",
  "Country": "generation:a81f...:key:country:"
}
```

The unprefixed manifests are stored in the `default` column family:

- `__xlsx2rocksdb:active-generations` maps each sheet to the generation readers must use, for example `{"Customers":"a81f...","Orders":"bf20..."}`.
- `__xlsx2rocksdb:generation-history` maps each sheet to its retained generation IDs in newest-first order, for example `{"Customers":["a81f...","previous..."]}`.

All selected sheets are written into the new generation without modifying the active generation. After the complete import succeeds, both manifests are replaced in one atomic RocksDB write batch. Unselected sheets retain their prior mappings. Generations older than `--keep-generations` are removed after publication; the history manifest is intended for retention and diagnostics, while normal reads must use the active manifest.

A consistent reader should:

1. Acquire one RocksDB snapshot.
2. Read `__xlsx2rocksdb:active-generations` from the `default` column family through that snapshot.
3. Form `generation:<uuid>:` for the requested sheet.
4. Read its metadata, key discovery, indexes, and rows through the same snapshot.
5. Release the snapshot.

If a sheet has no active-generation entry, readers may fall back to the unprefixed layout for legacy or mixed databases. RocksDB permits only one primary process; a concurrent server process must use a RocksDB secondary instance and call `tryCatchUpWithPrimary()` before starting a fresh request snapshot. See [Reading xlsx2rocksdb databases](doc/reader-guide.md) for a complete RocksJava implementation and operational guidance.

## Contributing and security

Contributions are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md) for the development and pull-request workflow. Report security vulnerabilities privately as described in [SECURITY.md](SECURITY.md). Changes are recorded in [CHANGELOG.md](CHANGELOG.md).

## License

This project is available under the [MIT License](LICENSE). Third-party dependencies remain subject to their respective licenses.

## Disclaimer

Parts of the software and documentation are created / edited using AI tools.
