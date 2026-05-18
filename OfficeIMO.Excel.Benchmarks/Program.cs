using OfficeIMO.Excel.Benchmarks;
using BenchmarkDotNet.Running;
using System.Globalization;

if (HasSwitch(args, "--help") || HasSwitch(args, "-h") || HasSwitch(args, "/?")) {
    WriteUsage();
    return;
}

if (IsCommand(args, "--snapshot", "snapshot")) {
    bool hasOutputPath = HasOutputPath(args);
    int rowCount = ParseRowCount(args, startIndex: hasOutputPath ? 2 : 1);
    string? websiteDataPath = ParseOptionValue(args, "--website-data", "--website-benchmarks");
    string? outputPathOverride = ParseOutputPath(args);
    string outputPath = ExcelBenchmarkSnapshotRunner.WriteSnapshot(
        outputPathOverride ?? BuildDefaultOutputPath("officeimo.excel.snapshot", rowCount),
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
    string? outputPathOverride = ParseOutputPath(args);
    string outputPath = ExcelWriteProfileRunner.WriteProfile(
        outputPathOverride ?? BuildDefaultOutputPath("officeimo.excel.write-profile", rowCount),
        rowCount);
    Console.WriteLine($"Excel write profile written to '{outputPath}'.");
    return;
}

if (IsCommand(args, "--profile-read", "profile-read", "read-profile")) {
    bool hasOutputPath = HasOutputPath(args);
    int rowCount = ParseRowCount(args, startIndex: hasOutputPath ? 2 : 1);
    int warmupIterations = ParsePositiveOption(args, "--warmup", "--warmups") ?? ExcelReadProfileRunner.DefaultWarmupIterations;
    int measuredIterations = ParsePositiveOption(args, "--iterations", "--measured-iterations", "--samples") ?? ExcelReadProfileRunner.DefaultMeasuredIterations;
    string? outputPathOverride = ParseOutputPath(args);
    string outputPath = ExcelReadProfileRunner.WriteProfile(
        outputPathOverride ?? BuildDefaultOutputPath("officeimo.excel.read-profile", rowCount),
        rowCount,
        warmupIterations,
        measuredIterations);
    Console.WriteLine($"Excel read profile written to '{outputPath}'.");
    return;
}

if (IsCommand(args, "--compare-libraries", "compare-libraries", "compare")) {
    bool hasOutputPath = HasOutputPath(args);
    int rowCount = ParseRowCount(args, startIndex: hasOutputPath ? 2 : 1);
    bool includeLegacyEpPlus = !HasSwitch(args, "--skip-legacy-epplus");
    string[] scenarioFilters = ParseOptionValues(args, "--scenario", "--scenarios");
    int warmupIterations = ParsePositiveOption(args, "--warmup", "--warmups") ?? ExcelLibraryComparisonRunner.DefaultWarmupIterations;
    int measuredIterations = ParsePositiveOption(args, "--iterations", "--measured-iterations", "--samples") ?? ExcelLibraryComparisonRunner.DefaultMeasuredIterations;
    string? outputPathOverride = ParseOutputPath(args);
    string outputPath = ExcelLibraryComparisonRunner.WriteComparison(
        outputPathOverride ?? BuildDefaultOutputPath("officeimo.excel.library-comparison", rowCount),
        rowCount,
        includeLegacyEpPlus,
        scenarioFilters,
        warmupIterations,
        measuredIterations);
    Console.WriteLine($"Excel library comparison written to '{outputPath}'.");
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);

static void WriteUsage() {
    Console.WriteLine("OfficeIMO.Excel benchmark helpers");
    Console.WriteLine();
    Console.WriteLine("Commands:");
    Console.WriteLine("  snapshot [output] [--rows N] [--website-data path]");
    Console.WriteLine("  write-profile [output] [--rows N]");
    Console.WriteLine("  read-profile [output] [--rows N] [--warmup N] [--iterations N]");
    Console.WriteLine("  compare [output] [--rows N] [--scenario name] [--skip-legacy-epplus] [--warmup N] [--iterations N]");
    Console.WriteLine();
    Console.WriteLine("Example:");
    Console.WriteLine("  compare .tmp\\officeimo.excel.library-comparison.json --rows 25000 --scenario write-dataset-tables --skip-legacy-epplus");
}

static bool IsCommand(string[] args, params string[] names)
    => args.Length >= 1 && names.Any(name => string.Equals(args[0], name, StringComparison.OrdinalIgnoreCase));

static bool HasOutputPath(string[] args)
    => args.Length >= 2 && !args[1].StartsWith("-", StringComparison.Ordinal);

static string? ParseOutputPath(string[] args)
    => ParseOptionValue(args, "--out", "--output", "--output-path")
       ?? (HasOutputPath(args) ? args[1] : null);

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

static string[] ParseOptionValues(string[] args, params string[] optionNames) {
    var values = new List<string>();
    for (int i = 0; i < args.Length; i++) {
        if (!optionNames.Any(name => string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))) {
            continue;
        }

        if (i + 1 >= args.Length || args[i + 1].StartsWith("-", StringComparison.Ordinal)) {
            throw new ArgumentException($"Missing value for {args[i]}.");
        }

        values.AddRange(args[i + 1]
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(value => value.Length > 0));
        i++;
    }

    return values.ToArray();
}

static int? ParsePositiveOption(string[] args, params string[] optionNames) {
    for (int i = 0; i < args.Length; i++) {
        if (!optionNames.Any(name => string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))) {
            continue;
        }

        if (i + 1 >= args.Length || args[i + 1].StartsWith("-", StringComparison.Ordinal)) {
            throw new ArgumentException($"Missing value for {args[i]}.");
        }

        string value = args[i + 1].Replace(",", string.Empty, StringComparison.Ordinal).Replace("_", string.Empty, StringComparison.Ordinal);
        if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) || parsed <= 0) {
            throw new ArgumentException($"{args[i]} must be a positive integer.");
        }

        return parsed;
    }

    return null;
}

static bool HasSwitch(string[] args, string optionName)
    => args.Any(arg => string.Equals(arg, optionName, StringComparison.OrdinalIgnoreCase));
