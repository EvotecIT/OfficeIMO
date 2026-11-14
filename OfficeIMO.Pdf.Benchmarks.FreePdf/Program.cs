using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Running;
using OfficeIMO.Pdf.Benchmarks.Comparisons;

ManualConfig config = ManualConfig
    .Create(DefaultConfig.Instance)
    .AddExporter(JsonExporter.Full);

BenchmarkSwitcher
    .FromAssembly(typeof(PdfSharpEngineGenerationBenchmarks).Assembly)
    .Run(args, config);

return 0;
