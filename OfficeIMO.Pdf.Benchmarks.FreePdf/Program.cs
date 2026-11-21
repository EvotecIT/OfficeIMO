using System.Linq;
using System.Reflection;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Running;
using OfficeIMO.Pdf.Benchmarks.Comparisons;

// BenchmarkDotNet rebuilds this project for out-of-process runs without the PdfSharpEngine property this build used.
// MSBuild reads environment variables as properties, so the child build compiles the same engine.
Assembly assembly = typeof(PdfSharpEngineGenerationBenchmarks).Assembly;
string? engine = assembly.GetCustomAttributes<AssemblyMetadataAttribute>()
    .FirstOrDefault(attribute => attribute.Key == "PdfSharpEngine")?.Value;
if (!string.IsNullOrEmpty(engine)) {
    System.Environment.SetEnvironmentVariable("PdfSharpEngine", engine);
}

ManualConfig config = ManualConfig
    .Create(DefaultConfig.Instance)
    .AddExporter(JsonExporter.Full);

BenchmarkSwitcher
    .FromAssembly(assembly)
    .Run(args, config);

return 0;
