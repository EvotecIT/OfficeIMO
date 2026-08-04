using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;
using OfficeIMO.Benchmarks;
using OfficeIMO.CSV.Benchmarks;

bool profileOfficeIMO = args.Length > 0 &&
    string.Equals(args[0], "--profile-markpflug65k-officeimo", StringComparison.OrdinalIgnoreCase);
bool profileSep = args.Length > 0 &&
    string.Equals(args[0], "--profile-markpflug65k-sep", StringComparison.OrdinalIgnoreCase);
bool profileSylvan = args.Length > 0 &&
    string.Equals(args[0], "--profile-markpflug65k-sylvan", StringComparison.OrdinalIgnoreCase);

if (profileOfficeIMO || profileSep || profileSylvan) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 100;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }

    var benchmark = new MarkPflug65KCsvBenchmarks();
    benchmark.Setup();
    Func<CsvReadObservation> run = profileOfficeIMO
        ? benchmark.OfficeIMO
        : profileSep
            ? benchmark.Sep
            : benchmark.Sylvan;
    string implementation = profileOfficeIMO
        ? "OfficeIMO"
        : profileSep
            ? "Sep"
            : "Sylvan";
    for (int index = 0; index < 3; index++) {
        run();
    }

    CsvReadObservation observation = default;
    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
    for (int index = 0; index < iterations; index++) {
        observation = run();
    }
    stopwatch.Stop();

    Console.WriteLine(
        $"Profiled {implementation} CSV {iterations} times in {stopwatch.Elapsed.TotalMilliseconds:F2} ms " +
        $"({stopwatch.Elapsed.TotalMilliseconds / iterations:F3} ms/iteration): {observation}.");
    return;
}

var config = ManualConfig
    .Create(DefaultConfig.Instance)
    .AddDiagnoser(MemoryDiagnoser.Default)
    .AddExporter(JsonExporter.Full)
    .WithSummaryStyle(SummaryStyle.Default.WithRatioStyle(RatioStyle.Percentage))
    .AddColumn(StatisticColumn.OperationsPerSecond);

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, config);
