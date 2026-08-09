using OfficeIMO.Excel.Benchmarks;
using OfficeIMO.Benchmarks;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using System.Globalization;
using System.Text.Json;

bool profileOfficeIMOXlsb = args.Length > 0 &&
    string.Equals(args[0], "--profile-markpflug65k-xlsb-officeimo", StringComparison.OrdinalIgnoreCase);
bool profileSylvanXlsb = args.Length > 0 &&
    string.Equals(args[0], "--profile-markpflug65k-xlsb-sylvan", StringComparison.OrdinalIgnoreCase);
bool profileOfficeIMOXls = args.Length > 0 &&
    string.Equals(args[0], "--profile-markpflug65k-xls-officeimo", StringComparison.OrdinalIgnoreCase);
bool profileSylvanXls = args.Length > 0 &&
    string.Equals(args[0], "--profile-markpflug65k-xls-sylvan", StringComparison.OrdinalIgnoreCase);
bool comparePairedXlsb = args.Length > 0 &&
    string.Equals(args[0], "--compare-markpflug65k-xlsb-paired", StringComparison.OrdinalIgnoreCase);
bool comparePairedXls = args.Length > 0 &&
    string.Equals(args[0], "--compare-markpflug65k-xls-paired", StringComparison.OrdinalIgnoreCase);
bool profileOfficeIMODataReaderWrite = args.Length > 0 &&
    string.Equals(args[0], "--profile-datareader-write-officeimo", StringComparison.OrdinalIgnoreCase);
bool profileLargeXlsxDataReaderWrite = args.Length > 0 &&
    string.Equals(args[0], "--profile-datareader-write-largexlsx", StringComparison.OrdinalIgnoreCase);
bool profileSpreadCheetahDataReaderWrite = args.Length > 0 &&
    string.Equals(args[0], "--profile-datareader-write-spreadcheetah", StringComparison.OrdinalIgnoreCase);
bool comparePairedDataReaderWrite = args.Length > 0 &&
    string.Equals(args[0], "--compare-datareader-write-paired", StringComparison.OrdinalIgnoreCase);

if (comparePairedDataReaderWrite) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 40;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    string affinity = ApplyProcessAffinity(args, argumentIndex: 2);
    string priority = ApplyProcessPriority(args, argumentIndex: 3);
    const int warmupIterations = 12;

    var benchmark = new ExcelDataReaderWriteBenchmarks { RowCount = 25_000 };
    benchmark.Setup();
    for (int index = 0; index < warmupIterations; index++) {
        benchmark.OfficeIMO();
        benchmark.SpreadCheetah();
    }

    var officeSamples = new double[iterations];
    var spreadSamples = new double[iterations];
    var pairedRatios = new double[iterations];
    for (int index = 0; index < iterations; index++) {
        double officeFirst;
        double officeSecond;
        double spreadFirst;
        double spreadSecond;
        if ((index & 1) == 0) {
            officeFirst = MeasureWriteMilliseconds(benchmark.OfficeIMO);
            spreadFirst = MeasureWriteMilliseconds(benchmark.SpreadCheetah);
            spreadSecond = MeasureWriteMilliseconds(benchmark.SpreadCheetah);
            officeSecond = MeasureWriteMilliseconds(benchmark.OfficeIMO);
        } else {
            spreadFirst = MeasureWriteMilliseconds(benchmark.SpreadCheetah);
            officeFirst = MeasureWriteMilliseconds(benchmark.OfficeIMO);
            officeSecond = MeasureWriteMilliseconds(benchmark.OfficeIMO);
            spreadSecond = MeasureWriteMilliseconds(benchmark.SpreadCheetah);
        }

        officeSamples[index] = (officeFirst + officeSecond) / 2d;
        spreadSamples[index] = (spreadFirst + spreadSecond) / 2d;
        pairedRatios[index] = officeSamples[index] / spreadSamples[index];
    }

    double officeMedian = Median(officeSamples);
    double spreadMedian = Median(spreadSamples);
    double ratioMedian = Median(pairedRatios);
    Console.WriteLine(
        $"Paired compact DataReader write ({warmupIterations} warmups, {iterations} ABBA samples, affinity {affinity}, priority {priority}): " +
        $"OfficeIMO median {officeMedian:F3} ms, SpreadCheetah median {spreadMedian:F3} ms, " +
        $"ratio of medians {officeMedian / spreadMedian:F4}, paired ratio median {ratioMedian:F4} " +
        $"(P25 {Percentile(pairedRatios, 0.25d):F4}, P75 {Percentile(pairedRatios, 0.75d):F4}).");
    return;
}

if (profileOfficeIMODataReaderWrite || profileLargeXlsxDataReaderWrite || profileSpreadCheetahDataReaderWrite) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 100;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    ApplyProcessAffinity(args, argumentIndex: 2);

    var benchmark = new ExcelDataReaderWriteBenchmarks { RowCount = 25_000 };
    benchmark.Setup();
    Func<int> write = profileOfficeIMODataReaderWrite
        ? benchmark.OfficeIMO
        : profileLargeXlsxDataReaderWrite ? benchmark.LargeXlsx : benchmark.SpreadCheetah;
    string implementation = profileOfficeIMODataReaderWrite
        ? "OfficeIMO"
        : profileLargeXlsxDataReaderWrite ? "LargeXlsx" : "SpreadCheetah";
    for (int index = 0; index < 3; index++) {
        write();
    }

    int outputBytes = 0;
    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
    for (int index = 0; index < iterations; index++) {
        outputBytes = write();
    }
    stopwatch.Stop();

    Console.WriteLine(
        $"Profiled {implementation} compact DataReader write {iterations} times in {stopwatch.Elapsed.TotalMilliseconds:F2} ms " +
        $"({stopwatch.Elapsed.TotalMilliseconds / iterations:F3} ms/iteration); last package {outputBytes:N0} bytes.");
    return;
}

if (comparePairedXlsb || comparePairedXls) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 20;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    string affinity = ApplyProcessAffinity(args, argumentIndex: 2);
    const int warmupIterations = 10;

    Func<ExcelReadObservation> runOfficeIMO;
    Func<ExcelReadObservation> runSylvan;
    string format;
    if (comparePairedXlsb) {
        var benchmark = new MarkPflug65KXlsbBenchmarks();
        benchmark.Setup();
        runOfficeIMO = benchmark.OfficeIMO;
        runSylvan = benchmark.Sylvan;
        format = "XLSB";
    } else {
        var benchmark = new MarkPflug65KXlsBenchmarks();
        benchmark.Setup();
        runOfficeIMO = benchmark.OfficeIMO;
        runSylvan = benchmark.Sylvan;
        format = "XLS";
    }
    for (int index = 0; index < warmupIterations; index++) {
        runOfficeIMO();
        runSylvan();
    }

    var officeSamples = new double[iterations];
    var sylvanSamples = new double[iterations];
    var pairedRatios = new double[iterations];
    for (int index = 0; index < iterations; index++) {
        ExcelReadObservation officeObservation;
        ExcelReadObservation sylvanObservation;
        if ((index & 1) == 0) {
            officeSamples[index] = MeasureMilliseconds(runOfficeIMO, out officeObservation);
            sylvanSamples[index] = MeasureMilliseconds(runSylvan, out sylvanObservation);
        } else {
            sylvanSamples[index] = MeasureMilliseconds(runSylvan, out sylvanObservation);
            officeSamples[index] = MeasureMilliseconds(runOfficeIMO, out officeObservation);
        }

        if (officeObservation != sylvanObservation) {
            throw new InvalidDataException(
                $"Paired {format} sample {index} produced different observations: OfficeIMO={officeObservation}; Sylvan={sylvanObservation}.");
        }
        pairedRatios[index] = officeSamples[index] / sylvanSamples[index];
    }

    double officeMedian = Median(officeSamples);
    double sylvanMedian = Median(sylvanSamples);
    double pairedRatioMedian = Median(pairedRatios);
    double pairedRatioP25 = Percentile(pairedRatios, 0.25d);
    double pairedRatioP75 = Percentile(pairedRatios, 0.75d);
    Console.WriteLine(
        $"Paired {format} comparison ({warmupIterations} warmups, {iterations} alternating samples, affinity {affinity}): " +
        $"OfficeIMO median {officeMedian:F3} ms, Sylvan median {sylvanMedian:F3} ms, " +
        $"ratio of medians {officeMedian / sylvanMedian:F4}, paired ratio median {pairedRatioMedian:F4} " +
        $"(P25 {pairedRatioP25:F4}, P75 {pairedRatioP75:F4}).");
    return;
}

if (profileOfficeIMOXlsb || profileSylvanXlsb || profileOfficeIMOXls || profileSylvanXls) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 100;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    ApplyProcessAffinity(args, argumentIndex: 2);

    bool isXlsb = profileOfficeIMOXlsb || profileSylvanXlsb;
    bool isOfficeIMO = profileOfficeIMOXlsb || profileOfficeIMOXls;
    Func<ExcelReadObservation> run;
    if (isXlsb) {
        var benchmark = new MarkPflug65KXlsbBenchmarks();
        benchmark.Setup();
        run = isOfficeIMO ? benchmark.OfficeIMO : benchmark.Sylvan;
    } else {
        var benchmark = new MarkPflug65KXlsBenchmarks();
        benchmark.Setup();
        run = isOfficeIMO ? benchmark.OfficeIMO : benchmark.Sylvan;
    }
    string implementation = isOfficeIMO ? "OfficeIMO" : "Sylvan";
    string format = isXlsb ? "XLSB" : "XLS";
    for (int index = 0; index < 3; index++) {
        run();
    }

    ExcelReadObservation observation = default;
    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
    for (int index = 0; index < iterations; index++) {
        observation = run();
    }
    stopwatch.Stop();

    Console.WriteLine(
        $"Profiled {implementation} {format} {iterations} times in {stopwatch.Elapsed.TotalMilliseconds:F2} ms " +
        $"({stopwatch.Elapsed.TotalMilliseconds / iterations:F3} ms/iteration): {observation}.");
    return;
}

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
    string? processorAffinity = ApplyProcessAffinityOption(args);
    string? outputPathOverride = ParseOutputPath(args);
    string outputPath = ExcelReadProfileRunner.WriteProfile(
        outputPathOverride ?? BuildDefaultOutputPath("officeimo.excel.read-profile", rowCount),
        rowCount,
        warmupIterations,
        measuredIterations,
        processorAffinity);
    Console.WriteLine($"Excel read profile written to '{outputPath}'.");
    return;
}

if (IsCommand(args, "--profile-chart", "profile-chart", "chart-profile")) {
    bool hasOutputPath = HasOutputPath(args);
    int rowCount = ParseRowCount(args, startIndex: hasOutputPath ? 2 : 1);
    int warmupIterations = ParsePositiveOption(args, "--warmup", "--warmups") ?? ExcelChartProfileRunner.DefaultWarmupIterations;
    int measuredIterations = ParsePositiveOption(args, "--iterations", "--measured-iterations", "--samples") ?? ExcelChartProfileRunner.DefaultMeasuredIterations;
    string? outputPathOverride = ParseOutputPath(args);
    string outputPath = ExcelChartProfileRunner.WriteProfile(
        outputPathOverride ?? BuildDefaultOutputPath("officeimo.excel.chart-profile", rowCount),
        rowCount,
        warmupIterations,
        measuredIterations);
    Console.WriteLine($"Excel chart profile written to '{outputPath}'.");
    return;
}

if (IsCommand(args, "--profile-realworld", "profile-realworld", "realworld-profile")) {
    bool hasOutputPath = HasOutputPath(args);
    int rowCount = ParseRowCount(args, startIndex: hasOutputPath ? 2 : 1);
    int warmupIterations = ParsePositiveOption(args, "--warmup", "--warmups") ?? ExcelRealWorldProfileRunner.DefaultWarmupIterations;
    int measuredIterations = ParsePositiveOption(args, "--iterations", "--measured-iterations", "--samples") ?? ExcelRealWorldProfileRunner.DefaultMeasuredIterations;
    string? outputPathOverride = ParseOutputPath(args);
    string outputPath = ExcelRealWorldProfileRunner.WriteProfile(
        outputPathOverride ?? BuildDefaultOutputPath("officeimo.excel.realworld-profile", rowCount),
        rowCount,
        warmupIterations,
        measuredIterations);
    Console.WriteLine($"Excel real-world profile written to '{outputPath}'.");
    return;
}

if (IsCommand(args, "--compare-libraries", "compare-libraries", "compare")) {
    bool hasOutputPath = HasOutputPath(args);
    int rowCount = ParseRowCount(args, startIndex: hasOutputPath ? 2 : 1);
    bool includeLegacyEpPlus = !HasSwitch(args, "--skip-legacy-epplus");
    string[] scenarioFilters = ParseOptionValues(args, "--scenario", "--scenarios");
    string[] libraryFilters = ParseOptionValues(args, "--library", "--libraries");
    int warmupIterations = ParsePositiveOption(args, "--warmup", "--warmups") ?? ExcelLibraryComparisonRunner.DefaultWarmupIterations;
    int measuredIterations = ParsePositiveOption(args, "--iterations", "--measured-iterations", "--samples") ?? ExcelLibraryComparisonRunner.DefaultMeasuredIterations;
    string? outputPathOverride = ParseOutputPath(args);
    string outputPath = ExcelLibraryComparisonRunner.WriteComparison(
        outputPathOverride ?? BuildDefaultOutputPath("officeimo.excel.library-comparison", rowCount),
        rowCount,
        includeLegacyEpPlus,
        scenarioFilters,
        warmupIterations,
        measuredIterations,
        libraryFilters);
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

if (IsCommand(args, "--anti-cheat-suite", "anti-cheat-suite", "robustness-suite", "variant-suite")) {
    bool hasOutputPath = HasOutputPath(args);
    string outputDirectory = ParseOptionValue(args, "--out-dir", "--output-dir", "--directory")
        ?? (hasOutputPath ? args[1] : Path.Combine("Docs", "benchmarks", "anti-cheat-current"));
    int[] rowCounts = ParseRowCountsOrDefault(args, startIndex: hasOutputPath ? 2 : 1, [100, 2500, 25000]);
    bool includeLegacyEpPlus = !HasSwitch(args, "--skip-legacy-epplus");
    bool includePackageProfile = !HasSwitch(args, "--skip-package-profile");
    string[] requestedScenarios = NormalizeScenarioFilters(ParseOptionValues(args, "--scenario", "--scenarios"));
    string[] scenarioFilters = requestedScenarios.Length == 0 ? GetAntiCheatScenarios() : requestedScenarios;
    string[] packageScenarioFilters = FilterPackageProfileScenarios(scenarioFilters);
    int warmupIterations = ParsePositiveOption(args, "--warmup", "--warmups") ?? ExcelLibraryComparisonRunner.DefaultWarmupIterations;
    int measuredIterations = ParsePositiveOption(args, "--iterations", "--measured-iterations", "--samples") ?? ExcelLibraryComparisonRunner.DefaultMeasuredIterations;

    Directory.CreateDirectory(outputDirectory);
    var artifacts = new List<ComparisonSuiteArtifact>();
    foreach (int rowCount in rowCounts) {
        string suffix = rowCount.ToString(CultureInfo.InvariantCulture);
        string comparisonPath = Path.Combine(outputDirectory, $"officeimo.excel.anti-cheat-speed-{suffix}.json");
        string writtenComparisonPath = ExcelLibraryComparisonRunner.WriteComparison(
            comparisonPath,
            rowCount,
            includeLegacyEpPlus,
            scenarioFilters,
            warmupIterations,
            measuredIterations);
        artifacts.Add(new ComparisonSuiteArtifact("speed-comparison", rowCount, writtenComparisonPath));
        Console.WriteLine($"Anti-cheat speed comparison written to '{writtenComparisonPath}'.");

        if (includePackageProfile && packageScenarioFilters.Length > 0) {
            string packagePath = Path.Combine(outputDirectory, $"officeimo.excel.anti-cheat-package-{suffix}.json");
            string writtenPackagePath = ExcelLibraryComparisonRunner.WritePackageProfile(
                packagePath,
                rowCount,
                packageScenarioFilters,
                warmupIterations,
                measuredIterations);
            artifacts.Add(new ComparisonSuiteArtifact("package-profile", rowCount, writtenPackagePath));
            Console.WriteLine($"Anti-cheat package profile written to '{writtenPackagePath}'.");
        }
    }

    var summary = ExcelComparisonSummaryWriter.WriteSummary(
        outputDirectory,
        artifacts.Select(artifact => new ExcelComparisonSummaryInput(artifact.Kind, artifact.RowCount, artifact.Path)),
        warmupIterations,
        measuredIterations);
    artifacts.Add(new ComparisonSuiteArtifact("summary-markdown", 0, summary.MarkdownPath));
    artifacts.Add(new ComparisonSuiteArtifact("summary-csv", 0, summary.CsvPath));
    artifacts.Add(new ComparisonSuiteArtifact("summary-json", 0, summary.JsonPath));
    Console.WriteLine($"Anti-cheat suite summary written to '{summary.MarkdownPath}'.");

    string manifestPath = Path.Combine(outputDirectory, "officeimo.excel.anti-cheat-suite-manifest.json");
    var manifest = new ComparisonSuiteManifest {
        GeneratedAtUtc = DateTime.UtcNow,
        Framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
        MachineName = Environment.MachineName,
        RowCounts = rowCounts,
        WarmupIterations = warmupIterations,
        MeasuredIterations = measuredIterations,
        IncludeLegacyEpPlus = includeLegacyEpPlus,
        IncludePackageProfile = includePackageProfile,
        IncludeDenseHelloWorld = false,
        ScenarioFilters = scenarioFilters,
        PackageScenarioFilters = packageScenarioFilters,
        DenseHelloWorldScenarios = [],
        Artifacts = artifacts
    };
    File.WriteAllText(manifestPath, JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine($"Anti-cheat suite manifest written to '{manifestPath}'.");
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
        artifacts.Select(artifact => new ExcelComparisonSummaryInput(artifact.Kind, artifact.RowCount, artifact.Path)),
        warmupIterations,
        measuredIterations);
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

var (priorityArgs, benchmarkPriority) = ExtractBenchmarkPriority(args);
if (benchmarkPriority != null) {
    BenchmarkProcessorAffinity.ApplyPriority(benchmarkPriority);
}
var (benchmarkArgs, affinityMasks) = ExtractAffinityMasks(priorityArgs);
var benchmarkConfig = ManualConfig
    .Create(DefaultConfig.Instance)
    .AddExporter(JsonExporter.Full);
for (int index = 0; index < affinityMasks.Length; index++) {
    IntPtr affinity = affinityMasks[index];
    Job job = Job.Default
        .WithAffinity(affinity)
        .WithId($"Affinity-{BenchmarkProcessorAffinity.Format(affinity)}");
    if (benchmarkPriority != null) {
        job = job.WithEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY", benchmarkPriority);
    }
    benchmarkConfig.AddJob(job);
}
BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(benchmarkArgs, benchmarkConfig);

static void WriteUsage() {
    Console.WriteLine("OfficeIMO.Excel benchmark helpers");
    Console.WriteLine();
    Console.WriteLine("Commands:");
    Console.WriteLine("  snapshot [output] [--rows N] [--website-data path]");
    Console.WriteLine("  write-profile [output] [--rows N]");
    Console.WriteLine("  read-profile [output] [--rows N] [--warmup N] [--iterations N] [--affinity mask]");
    Console.WriteLine("  chart-profile [output] [--rows N] [--warmup N] [--iterations N]");
    Console.WriteLine("  compare [output] [--rows N] [--scenario name] [--library name] [--skip-legacy-epplus] [--warmup N] [--iterations N]");
    Console.WriteLine("  package-profile [output] [--rows N] [--scenario name] [--warmup N] [--iterations N]");
    Console.WriteLine("  --profile-datareader-write-officeimo [iterations] [affinity-mask]");
    Console.WriteLine("  --profile-datareader-write-largexlsx [iterations] [affinity-mask]");
    Console.WriteLine("  --profile-datareader-write-spreadcheetah [iterations] [affinity-mask]");
    Console.WriteLine("  --compare-datareader-write-paired [iterations] [affinity-mask] [priority]");
    Console.WriteLine("  anti-cheat-suite [output-dir] [--row-set 100,2500,25000] [--scenario name] [--skip-legacy-epplus] [--skip-package-profile] [--warmup N] [--iterations N]");
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

    if (startIndex < args.Length && !args[startIndex].StartsWith("-", StringComparison.Ordinal)) {
        string value = args[startIndex].Replace(",", string.Empty, StringComparison.Ordinal).Replace("_", string.Empty, StringComparison.Ordinal);
        if (int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int rowCount)
            && rowCount > 0) {
            return rowCount;
        }
    }

    for (int i = startIndex; i < args.Length; i++) {
        string arg = args[i];
        if (arg.StartsWith("-", StringComparison.Ordinal)) {
            if (OptionConsumesValue(arg)) {
                i++;
            }

            continue;
        }

        string value = arg.Replace(",", string.Empty, StringComparison.Ordinal).Replace("_", string.Empty, StringComparison.Ordinal);
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int rowCount)
            && rowCount > 0) {
            return rowCount;
        }
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

static int[] ParseRowCountsOrDefault(string[] args, int startIndex, int[] defaultRowCounts) {
    if (!HasAnyOption(args, "--row-set", "--rows", "--row-counts")) {
        return defaultRowCounts
            .Where(rowCount => rowCount > 0)
            .Distinct()
            .Order()
            .ToArray();
    }

    return ParseRowCounts(args, startIndex);
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

static string? ApplyProcessAffinityOption(string[] args) {
    string? value = ParseOptionValue(args, "--affinity", "--processor-affinity");
    return value is null ? null : BenchmarkProcessorAffinity.Apply(value);
}

static (string[] Arguments, IntPtr[] Masks) ExtractAffinityMasks(string[] arguments) {
    int optionIndex = Array.FindIndex(
        arguments,
        static argument => string.Equals(argument, "--affinityMasks", StringComparison.OrdinalIgnoreCase));
    if (optionIndex < 0) {
        return (arguments, Array.Empty<IntPtr>());
    }
    if (optionIndex + 1 >= arguments.Length) {
        throw new ArgumentException("--affinityMasks requires a comma-separated list of positive processor-affinity masks.");
    }

    IntPtr[] masks = BenchmarkProcessorAffinity.ParseList(arguments[optionIndex + 1]);
    var forwarded = new string[arguments.Length - 2];
    if (optionIndex > 0) {
        Array.Copy(arguments, 0, forwarded, 0, optionIndex);
    }
    if (optionIndex + 2 < arguments.Length) {
        Array.Copy(
            arguments,
            optionIndex + 2,
            forwarded,
            optionIndex,
            arguments.Length - optionIndex - 2);
    }
    return (forwarded, masks);
}

static (string[] Arguments, string? Priority) ExtractBenchmarkPriority(string[] arguments) {
    int optionIndex = Array.FindIndex(
        arguments,
        static argument => string.Equals(argument, "--priority", StringComparison.OrdinalIgnoreCase));
    if (optionIndex < 0) {
        return (arguments, null);
    }
    if (optionIndex + 1 >= arguments.Length) {
        throw new ArgumentException("--priority requires Idle, BelowNormal, Normal, AboveNormal, or High.");
    }

    string priority = arguments[optionIndex + 1];
    var forwarded = new string[arguments.Length - 2];
    if (optionIndex > 0) {
        Array.Copy(arguments, 0, forwarded, 0, optionIndex);
    }
    if (optionIndex + 2 < arguments.Length) {
        Array.Copy(arguments, optionIndex + 2, forwarded, optionIndex, arguments.Length - optionIndex - 2);
    }
    return (forwarded, priority);
}

static bool HasSwitch(string[] args, string optionName)
    => args.Any(arg => string.Equals(arg, optionName, StringComparison.OrdinalIgnoreCase));

static bool HasAnyOption(string[] args, params string[] optionNames)
    => args.Any(arg => optionNames.Any(option => string.Equals(arg, option, StringComparison.OrdinalIgnoreCase)));

static bool OptionConsumesValue(string option)
    => string.Equals(option, "--out", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--output", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--output-path", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--out-dir", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--output-dir", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--directory", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--website-data", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--website-benchmarks", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--scenario", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--scenarios", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--library", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--libraries", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--warmup", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--warmups", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--iterations", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--measured-iterations", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--samples", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--rows", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--row-count", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--row-set", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--row-counts", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--affinity", StringComparison.OrdinalIgnoreCase)
       || string.Equals(option, "--processor-affinity", StringComparison.OrdinalIgnoreCase);

static string[] FilterPackageProfileScenarios(IReadOnlyCollection<string> scenarioFilters) {
    if (scenarioFilters.Count == 0) {
        return [];
    }

    var packageScenarios = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "write-bulk-report",
        "write-dataset-tables",
        "write-dataset-tables-autofit",
        "write-dataset-sparse-tables",
        "write-dataset-sparse-direct-export",
        "write-datatable-direct",
        "write-datatable-table-direct",
        "write-datareader-table",
        "write-datareader-table-autofit",
        "write-datareader-plain",
        "write-datareader-direct-package",
        "write-datareader-compact-package",
        "write-cellvalues-rectangle-direct",
        "write-cellvalues-headerless-rectangle-direct",
        "write-cellvalue-strings",
        "write-cellvalue-strings-repeated",
        "write-cellvalue-strings-distinct",
        "write-cellvalue-empty-strings",
        "write-cellvalue-numbers",
        "write-cellvalue-scalars",
        "write-cellvalue-temporal",
        "write-cellvalue-object-mixed",
        "write-cellvalue-object-sparse",
        "write-cellvalue-object-sparse-batch",
        "write-cellformula",
        "write-insertobjects-direct",
        "write-typed-rows-direct-package",
        "write-typed-rows-compact-package",
        "write-insertobjects-autofitcolumnsfor-direct",
        "write-insertobjects-partial-autofitcolumnsfor-direct",
        "write-insertobjects-flat-dictionaries-direct",
        "write-insertobjects-legacy-dictionaries-direct",
        "write-powershell-mixed-objects-direct",
        "write-powershell-psobject-mixed-direct",
        "write-powershell-psobject-wide-direct",
        "write-fluent-rowsfrom-direct",
        "append-plain-rows",
        "autofit-existing",
        "report-workbook",
        "report-workbook-core",
        "report-workbook-datatable",
        "report-workbook-datatable-core",
        "realworld-report-all-in-one",
        "realworld-report-core",
        "realworld-freeze-panes",
        "realworld-autofilter",
        "realworld-conditional-formatting",
        "realworld-data-validation",
        "realworld-charts",
        "realworld-pivot-table",
        "realworld-report-no-autofit",
        "realworld-report-chart-first",
        "realworld-report-shuffled-columns",
        "realworld-report-extra-column",
        "realworld-report-post-mutation",
        "write-text-heavy-default"
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

static string[] GetAntiCheatScenarios()
    => [
        "realworld-report-all-in-one",
        "realworld-report-no-autofit",
        "realworld-report-chart-first",
        "realworld-report-shuffled-columns",
        "realworld-report-extra-column",
        "realworld-report-post-mutation"
    ];

static double MeasureMilliseconds(
    Func<ExcelReadObservation> operation,
    out ExcelReadObservation observation) {
    long started = System.Diagnostics.Stopwatch.GetTimestamp();
    observation = operation();
    return System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds;
}

static string ApplyProcessAffinity(string[] arguments, int argumentIndex)
    => arguments.Length <= argumentIndex
        ? "unchanged"
        : BenchmarkProcessorAffinity.Apply(arguments[argumentIndex]);

static string ApplyProcessPriority(string[] arguments, int argumentIndex) {
    if (arguments.Length <= argumentIndex ||
        string.Equals(arguments[argumentIndex], "unchanged", StringComparison.OrdinalIgnoreCase)) {
        return System.Diagnostics.Process.GetCurrentProcess().PriorityClass.ToString();
    }

    return BenchmarkProcessorAffinity.ApplyPriority(arguments[argumentIndex]);
}

static double MeasureWriteMilliseconds(Func<int> operation) {
    long started = System.Diagnostics.Stopwatch.GetTimestamp();
    _ = operation();
    return System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds;
}

static double Median(double[] samples) {
    Array.Sort(samples);
    int middle = samples.Length / 2;
    return (samples.Length & 1) == 0
        ? (samples[middle - 1] + samples[middle]) / 2d
        : samples[middle];
}

static double Percentile(double[] samples, double percentile) {
    if (samples.Length == 0) {
        throw new ArgumentException("At least one sample is required.", nameof(samples));
    }
    if (percentile < 0d || percentile > 1d) {
        throw new ArgumentOutOfRangeException(nameof(percentile));
    }

    Array.Sort(samples);
    double position = (samples.Length - 1) * percentile;
    int lower = (int)position;
    int upper = Math.Min(lower + 1, samples.Length - 1);
    double fraction = position - lower;
    return samples[lower] + (samples[upper] - samples[lower]) * fraction;
}

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
