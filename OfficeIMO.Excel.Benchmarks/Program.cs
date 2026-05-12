using OfficeIMO.Excel.Benchmarks;
using BenchmarkDotNet.Running;
using System.Globalization;

if (IsCommand(args, "--snapshot", "snapshot")) {
    bool hasOutputPath = HasOutputPath(args);
    int rowCount = ParseRowCount(args, startIndex: hasOutputPath ? 2 : 1);
    string? websiteDataPath = ParseOptionValue(args, "--website-data", "--website-benchmarks");
    string outputPath = ExcelBenchmarkSnapshotRunner.WriteSnapshot(
        hasOutputPath ? args[1] : BuildDefaultOutputPath("officeimo.excel.snapshot", rowCount),
        rowCount,
        websiteDataPath);
    Console.WriteLine($"Excel benchmark snapshot written to '{outputPath}'.");
    if (!string.IsNullOrWhiteSpace(websiteDataPath)) {
        Console.WriteLine($"Website benchmark data updated at '{websiteDataPath}'.");
    }
    return;
}

if (IsCommand(args, "--profile-write", "profile-write", "write-profile")) {
    bool hasOutputPath = HasOutputPath(args);
    int rowCount = ParseRowCount(args, startIndex: hasOutputPath ? 2 : 1);
    string outputPath = ExcelWriteProfileRunner.WriteProfile(
        hasOutputPath ? args[1] : BuildDefaultOutputPath("officeimo.excel.write-profile", rowCount),
        rowCount);
    Console.WriteLine($"Excel write profile written to '{outputPath}'.");
    return;
}

if (IsCommand(args, "--profile-read", "profile-read", "read-profile")) {
    bool hasOutputPath = HasOutputPath(args);
    int rowCount = ParseRowCount(args, startIndex: hasOutputPath ? 2 : 1);
    string outputPath = ExcelReadProfileRunner.WriteProfile(
        hasOutputPath ? args[1] : BuildDefaultOutputPath("officeimo.excel.read-profile", rowCount),
        rowCount);
    Console.WriteLine($"Excel read profile written to '{outputPath}'.");
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);

static bool IsCommand(string[] args, params string[] names)
    => args.Length >= 1 && names.Any(name => string.Equals(args[0], name, StringComparison.OrdinalIgnoreCase));

static bool HasOutputPath(string[] args)
    => args.Length >= 2 && !args[1].StartsWith("-", StringComparison.Ordinal);

static string BuildDefaultOutputPath(string baseName, int rowCount) {
    string suffix = rowCount == 2500 ? string.Empty : "-" + rowCount.ToString(CultureInfo.InvariantCulture);
    return Path.Combine("Docs", "benchmarks", baseName + suffix + ".json");
}

static int ParseRowCount(string[] args, int startIndex) {
    const int defaultRowCount = 2500;

    for (int i = startIndex; i < args.Length; i++) {
        if (!string.Equals(args[i], "--rows", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(args[i], "--row-count", StringComparison.OrdinalIgnoreCase)) {
            continue;
        }

        if (i + 1 >= args.Length) {
            throw new ArgumentException("Missing value for --rows.");
        }

        string value = args[i + 1].Replace(",", string.Empty, StringComparison.Ordinal).Replace("_", string.Empty, StringComparison.Ordinal);
        if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int rowCount)
            || rowCount <= 0) {
            throw new ArgumentException("--rows must be a positive integer.");
        }

        return rowCount;
    }

    return defaultRowCount;
}

static string? ParseOptionValue(string[] args, params string[] optionNames) {
    for (int i = 0; i < args.Length; i++) {
        if (!optionNames.Any(name => string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))) {
            continue;
        }

        if (i + 1 >= args.Length || args[i + 1].StartsWith("-", StringComparison.Ordinal)) {
            throw new ArgumentException($"Missing value for {args[i]}.");
        }

        return args[i + 1];
    }

    return null;
}
