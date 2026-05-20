using OfficeIMO.Excel.Benchmarks;
using BenchmarkDotNet.Running;
using System.Globalization;
using System.Text.Json;

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

if (IsCommand(args, "--package-profile", "package-profile", "profile-package")) {
    bool hasOutputPath = HasOutputPath(args);
    int rowCount = ParseRowCount(args, startIndex: hasOutputPath ? 2 : 1);
    string[] scenarioFilters = ParseOptionValues(args, "--scenario", "--scenarios");
    int warmupIterations = ParsePositiveOption(args, "--warmup", "--warmups") ?? ExcelLibraryComparisonRunner.DefaultWarmupIterations;
    int measuredIterations = ParsePositiveOption(args, "--iterations", "--measured-iterations", "--samples") ?? ExcelLibraryComparisonRunner.DefaultMeasuredIterations;
    string? outputPathOverride = ParseOutputPath(args);
    string outputPath = ExcelLibraryComparisonRunner.WritePackageProfile(
        outputPathOverride ?? BuildDefaultOutputPath("officeimo.excel.package-profile", rowCount),
        rowCount,
        scenarioFilters,
        warmupIterations,
        measuredIterations);
    Console.WriteLine($"Excel package profile written to '{outputPath}'.");
    return;
}

if (IsCommand(args, "--comparison-suite", "comparison-suite", "--competitive-suite", "competitive-suite", "suite")) {
    bool hasOutputPath = HasOutputPath(args);
    string outputDirectory = ParseOptionValue(args, "--out-dir", "--output-dir", "--directory")
        ?? (hasOutputPath ? args[1] : Path.Combine("Docs", "benchmarks"));
    int[] rowCounts = ParseRowCounts(args, startIndex: hasOutputPath ? 2 : 1);
    bool includeLegacyEpPlus = !HasSwitch(args, "--skip-legacy-epplus");
    bool includePackageProfile = !HasSwitch(args, "--skip-package-profile");
    bool includeDenseHelloWorld = !HasSwitch(args, "--skip-dense-helloworld") && !HasSwitch(args, "--skip-miniexcel-helloworld");
    string[] scenarioFilters = NormalizeScenarioFilters(ParseOptionValues(args, "--scenario", "--scenarios"));
    string[] packageScenarioFilters = FilterPackageProfileScenarios(scenarioFilters);
    string[] helloWorldScenarios = GetDenseHelloWorldScenarios();
    bool runHelloWorldSeparately = includeDenseHelloWorld && scenarioFilters.Length == 0;
    int warmupIterations = ParsePositiveOption(args, "--warmup", "--warmups") ?? ExcelLibraryComparisonRunner.DefaultWarmupIterations;
    int measuredIterations = ParsePositiveOption(args, "--iterations", "--measured-iterations", "--samples") ?? ExcelLibraryComparisonRunner.DefaultMeasuredIterations;

    Directory.CreateDirectory(outputDirectory);
    var artifacts = new List<ComparisonSuiteArtifact>();
    foreach (int rowCount in rowCounts) {
        string suffix = rowCount.ToString(CultureInfo.InvariantCulture);
        string comparisonPath = Path.Combine(outputDirectory, $"officeimo.excel.comparison-speed-{suffix}.json");
        string writtenComparisonPath = ExcelLibraryComparisonRunner.WriteComparison(
            comparisonPath,
            rowCount,
            includeLegacyEpPlus,
            scenarioFilters,
            warmupIterations,
            measuredIterations);
        artifacts.Add(new ComparisonSuiteArtifact("speed-comparison", rowCount, writtenComparisonPath));
        Console.WriteLine($"Suite speed comparison written to '{writtenComparisonPath}'.");

        if (includePackageProfile && (scenarioFilters.Length == 0 || packageScenarioFilters.Length > 0)) {
            string packagePath = Path.Combine(outputDirectory, $"officeimo.excel.comparison-package-{suffix}.json");
            string writtenPackagePath = ExcelLibraryComparisonRunner.WritePackageProfile(
                packagePath,
                rowCount,
                packageScenarioFilters,
                warmupIterations,
                measuredIterations);
            artifacts.Add(new ComparisonSuiteArtifact("package-profile", rowCount, writtenPackagePath));
            Console.WriteLine($"Suite package profile written to '{writtenPackagePath}'.");
        } else if (includePackageProfile) {
            Console.WriteLine("Package profile skipped because the requested scenario filter only contains read-only scenarios.");
        }

        if (runHelloWorldSeparately) {
            string helloWorldPath = Path.Combine(outputDirectory, $"officeimo.excel.comparison-dense-helloworld-{suffix}.json");
            string writtenHelloWorldPath = ExcelLibraryComparisonRunner.WriteComparison(
                helloWorldPath,
                rowCount,
                includeLegacyEpPlus: false,
                helloWorldScenarios,
                warmupIterations,
                measuredIterations);
            artifacts.Add(new ComparisonSuiteArtifact("dense-helloworld-comparison", rowCount, writtenHelloWorldPath));
            Console.WriteLine($"Dense HelloWorld comparison written to '{writtenHelloWorldPath}'.");
        }
    }

    var summary = ExcelComparisonSummaryWriter.WriteSummary(
        outputDirectory,
        artifacts.Select(artifact => new ExcelComparisonSummaryInput(artifact.Kind, artifact.RowCount, artifact.Path)));
    artifacts.Add(new ComparisonSuiteArtifact("summary-markdown", 0, summary.MarkdownPath));
    artifacts.Add(new ComparisonSuiteArtifact("summary-csv", 0, summary.CsvPath));
    artifacts.Add(new ComparisonSuiteArtifact("summary-json", 0, summary.JsonPath));
    Console.WriteLine($"Comparison suite summary written to '{summary.MarkdownPath}'.");

    string manifestPath = Path.Combine(outputDirectory, "officeimo.excel.comparison-suite-manifest.json");
    var manifest = new ComparisonSuiteManifest {
        GeneratedAtUtc = DateTime.UtcNow,
        Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
        MachineName = Environment.MachineName,
        RowCounts = rowCounts,
        WarmupIterations = warmupIterations,
        MeasuredIterations = measuredIterations,
        IncludeLegacyEpPlus = includeLegacyEpPlus,
        IncludePackageProfile = includePackageProfile,
        IncludeDenseHelloWorld = runHelloWorldSeparately,
        ScenarioFilters = scenarioFilters,
        PackageScenarioFilters = packageScenarioFilters,
        DenseHelloWorldScenarios = runHelloWorldSeparately ? helloWorldScenarios : [],
        Artifacts = artifacts
    };
    File.WriteAllText(manifestPath, JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine($"Comparison suite manifest written to '{manifestPath}'.");
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
    Console.WriteLine("  package-profile [output] [--rows N] [--scenario name] [--warmup N] [--iterations N]");
    Console.WriteLine("  comparison-suite [output-dir] [--row-set 2500,25000] [--scenario name] [--skip-legacy-epplus] [--skip-package-profile] [--skip-dense-helloworld] [--warmup N] [--iterations N]");
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

static int[] ParseRowCounts(string[] args, int startIndex) {
    var rowCounts = new List<int>();
    for (int i = startIndex; i < args.Length; i++) {
        if (!string.Equals(args[i], "--row-set", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(args[i], "--rows", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(args[i], "--row-counts", StringComparison.OrdinalIgnoreCase)) {
            continue;
        }

        if (i + 1 >= args.Length || args[i + 1].StartsWith("-", StringComparison.Ordinal)) {
            throw new ArgumentException($"Missing value for {args[i]}.");
        }

        foreach (string part in args[i + 1].Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) {
            string value = part.Replace("_", string.Empty, StringComparison.Ordinal);
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) || parsed <= 0) {
                throw new ArgumentException($"{args[i]} must contain positive integers separated by commas, for example 2500,25000.");
            }

            if (!rowCounts.Contains(parsed)) {
                rowCounts.Add(parsed);
            }
        }

        i++;
    }

    if (rowCounts.Count == 0) {
        rowCounts.Add(2500);
    }

    rowCounts.Sort();
    return rowCounts.ToArray();
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

static string[] FilterPackageProfileScenarios(IReadOnlyCollection<string> scenarioFilters) {
    if (scenarioFilters.Count == 0) {
        return [];
    }

    var packageScenarios = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "write-bulk-report",
        "write-dataset-tables",
        "write-dataset-tables-autofit",
        "write-datatable-direct",
        "write-datatable-table-direct",
        "write-datareader-table",
        "write-cellvalues-rectangle-direct",
        "write-insertobjects-direct",
        "write-fluent-rowsfrom-direct",
        "append-plain-rows",
        "autofit-existing",
        "large-shared-strings"
    };

    return scenarioFilters
        .Where(packageScenarios.Contains)
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToArray();
}

static string[] NormalizeScenarioFilters(string[] scenarioFilters) {
    if (scenarioFilters.Length == 0) {
        return scenarioFilters;
    }

    return scenarioFilters
        .Select(scenario => scenario.Equals("miniexcel-helloworld-read-range", StringComparison.OrdinalIgnoreCase)
            ? "dense-helloworld-read-range"
            : scenario.Equals("miniexcel-helloworld-read-stream", StringComparison.OrdinalIgnoreCase)
                ? "dense-helloworld-read-stream"
                : scenario)
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToArray();
}

static string[] GetDenseHelloWorldScenarios()
    => [
        "dense-helloworld-read-range",
        "dense-helloworld-read-stream"
    ];

internal sealed class ComparisonSuiteManifest {
    public DateTime GeneratedAtUtc { get; init; }
    public string Framework { get; init; } = string.Empty;
    public string MachineName { get; init; } = string.Empty;
    public int[] RowCounts { get; init; } = [];
    public int WarmupIterations { get; init; }
    public int MeasuredIterations { get; init; }
    public bool IncludeLegacyEpPlus { get; init; }
    public bool IncludePackageProfile { get; init; }
    public bool IncludeDenseHelloWorld { get; init; }
    public string[] ScenarioFilters { get; init; } = [];
    public string[] PackageScenarioFilters { get; init; } = [];
    public string[] DenseHelloWorldScenarios { get; init; } = [];
    public List<ComparisonSuiteArtifact> Artifacts { get; init; } = [];
}

internal sealed record ComparisonSuiteArtifact(string Kind, int RowCount, string Path);
