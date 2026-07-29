using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Order;
using BenchmarkDotNet.Running;

namespace OfficeIMO.Tabular.Benchmarks;

internal static class Program {
    public static void Main(string[] args) {
        if (args.Contains("--validate", StringComparer.OrdinalIgnoreCase)) {
            BenchmarkValidation.Run();
            return;
        }

        bool quick = args.Contains("--quick", StringComparer.OrdinalIgnoreCase);
        string artifactsPath = ReadArgument(args, "--artifacts")
            ?? Path.Combine(Environment.CurrentDirectory, "BenchmarkDotNet.Artifacts");
        string[] benchmarkArguments = RemoveRunnerArguments(args);
        if (benchmarkArguments.Length == 0) {
            benchmarkArguments = new[] { "--filter", "*" };
        }

        FixtureData.EnsureAuthentic();
        FixtureData.WriteProvenance(artifactsPath);

        Job job = quick
            ? Job.ShortRun.WithId("quick")
            : Job.Default.WithId("full");
        ManualConfig config = ManualConfig.Create(DefaultConfig.Instance)
            .WithArtifactsPath(artifactsPath)
            .AddJob(job)
            .AddDiagnoser(MemoryDiagnoser.Default)
            .AddExporter(JsonExporter.FullCompressed)
            .AddColumn(StatisticColumn.OperationsPerSecond)
            .WithOrderer(new DefaultOrderer(SummaryOrderPolicy.Declared));

        BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(benchmarkArguments, config);
    }

    private static string? ReadArgument(IReadOnlyList<string> args, string name) {
        for (int index = 0; index < args.Count - 1; index++) {
            if (string.Equals(args[index], name, StringComparison.OrdinalIgnoreCase)) {
                return args[index + 1];
            }
        }

        return null;
    }

    private static string[] RemoveRunnerArguments(IReadOnlyList<string> args) {
        var filtered = new List<string>();
        for (int index = 0; index < args.Count; index++) {
            if (string.Equals(args[index], "--quick", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }
            if (string.Equals(args[index], "--artifacts", StringComparison.OrdinalIgnoreCase)) {
                index++;
                continue;
            }

            filtered.Add(args[index]);
        }

        return filtered.ToArray();
    }
}
