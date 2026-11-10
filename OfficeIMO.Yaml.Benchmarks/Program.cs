using BenchmarkDotNet.Running;
using OfficeIMO.Yaml.Benchmarks.Comparison;
using OfficeIMO.Yaml.Benchmarks.Workloads;

if (args.Length > 0 && string.Equals(args[0], "--conformance", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = YamlConformanceComparison.Run(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--validate-workloads", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = YamlWorkloadCheck.Run();
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(YamlConformanceComparison).Assembly).Run(args);
