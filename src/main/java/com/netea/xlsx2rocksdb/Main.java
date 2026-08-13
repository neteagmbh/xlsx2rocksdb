package com.netea.xlsx2rocksdb;

import picocli.CommandLine;

public final class Main {
    private Main() {
    }

    public static void main(String[] args) {
        int exitCode = new CommandLine(new ImportCommand()).execute(args);
        System.exit(exitCode);
    }
}
