using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using OfficeIMO.Pdf.Benchmarks.Comparisons;

// This optional host reuses the production benchmark and validation helpers.
// Only the Core/Pdf assembly references differ between the two child processes.
long affinity = long.Parse(Environment.GetEnvironmentVariable("PDF_BENCHMARK_AFFINITY") ?? "65535", CultureInfo.InvariantCulture);
Job settings = Environment.GetEnvironmentVariable("PDF_BENCHMARK_RUN") switch {
    "dry" => Job.Dry,
    "short" => Job.ShortRun,
    "full" or null => Job.Default,
    _ => throw new ArgumentException("PDF_BENCHMARK_RUN must be full, short, or dry.")
};
var config = ManualConfig.Create(DefaultConfig.Instance).AddExporter(JsonExporter.Full);
config.AddJob(settings.WithAffinity(new IntPtr(affinity)).WithId("baseline").AsBaseline()
    .WithArguments(new Argument[] { new MsBuildArgument("/p:UseBaseline=true") }));
config.AddJob(settings.WithAffinity(new IntPtr(affinity)).WithId("current"));
BenchmarkSwitcher.FromTypes(new[] { typeof(PdfNativeOperationsBenchmarks) }).Run(args, config);
