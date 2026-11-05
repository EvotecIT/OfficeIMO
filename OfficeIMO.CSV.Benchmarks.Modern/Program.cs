using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Running;
using OfficeIMO.CSV.Benchmarks;

if (args.Length == 1 && args[0] == "--validate") {
    new MarkPflug65KModernCsvBenchmarks().Setup();
    Console.WriteLine("PASS | All four readers produce the complete expected owned-string observation.");
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(MarkPflug65KModernCsvBenchmarks).Assembly).Run(args,
    ManualConfig.Create(DefaultConfig.Instance).AddExporter(JsonExporter.Full));
