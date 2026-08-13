# Contributing to xlsx2rocksdb

Contributions are welcome through GitHub issues and pull requests.

## Development setup

Requirements:

- JDK 17 or newer
- Maven 3.9 or newer

Build and run the test suite:

```bash
mvn clean verify
```

The executable JAR is written to `target/xlsx2rocksdb.jar`. Tests create their own temporary XLSX files and RocksDB databases; local `testdata/` content is not required and is intentionally excluded from the repository.

## Pull requests

- Keep changes focused and include tests for behavior changes.
- Update the README or `doc/reader-guide.md` when the CLI or storage contract changes.
- Run `mvn clean verify` before submitting.
- Do not commit real workbooks, RocksDB directories, credentials, or confidential data.
- By contributing, you agree that your contribution is licensed under the MIT License.

For substantial behavior or storage-format changes, open an issue first so compatibility and migration requirements can be discussed.

## Release process

1. Choose the next Semantic Versioning number.
2. Update the version in `pom.xml`, the Picocli version string in `ImportCommand`, and `CHANGELOG.md`.
3. Run `mvn clean verify` and confirm `java -jar target/xlsx2rocksdb.jar --version`.
4. Merge the release commit to `main` and create a matching signed tag such as `v0.1.0`.
5. Push the tag. The release workflow tests the project and publishes the executable JAR plus its SHA-256 checksum to GitHub Releases.
