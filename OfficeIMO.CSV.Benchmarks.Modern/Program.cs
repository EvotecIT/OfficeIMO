using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Running;
using OfficeIMO.CSV.Benchmarks;

if (args.Length == 1 && args[0] == "--validate") {
    new MarkPflug65KModernCsvBenchmarks().Setup();
    new CsvGeneratedModernReadBenchmarks().Setup();
    Console.WriteLine("PASS | All five sales-reader configurations and all four generated-data readers produce the expected values.");
    return;
}

if (args.Length == 2 && args[0] == "--profile-interned-csv") {
    int count = int.Parse(args[1], System.Globalization.CultureInfo.InvariantCulture);
    ArgumentOutOfRangeException.ThrowIfNegativeOrZero(count);
    var benchmark = new MarkPflug65KModernCsvBenchmarks();
    benchmark.Setup();
    for (int iteration = 0; iteration < count; iteration++) _ = benchmark.OfficeIMO_InternStrings();
    Console.WriteLine($"PASS | Validated {count} complete interned CSV traversals.");
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(MarkPflug65KModernCsvBenchmarks).Assembly).Run(args,
    ManualConfig.Create(DefaultConfig.Instance).AddExporter(JsonExporter.Full));
