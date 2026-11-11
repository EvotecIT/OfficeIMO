using BenchmarkDotNet.Running;
using OfficeIMO.Toml.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--validate-workloads", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = TomlWorkloadCheck.Run();
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(TomlParseBenchmarks).Assembly).Run(args);
