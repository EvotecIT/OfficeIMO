using OfficeIMO.Yaml.Benchmarks;
using BenchmarkDotNet.Running;
using OfficeIMO.Yaml.Benchmarks.Comparison;
using OfficeIMO.Yaml.Benchmarks.LargeCorpus;
using OfficeIMO.Yaml.Benchmarks.Workloads;

if (args.Length > 0 && string.Equals(args[0], "--conformance", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = YamlConformanceComparison.Run(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--large-corpus", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = YamlLargeCorpusRun.Run(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--merge-corpus", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = YamlMergeCorpusRun.Run(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--mutations", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = YamlMutationRun.Run(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--serializer-timings", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = YamlSerializerTimings.Run(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--validate-workloads", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = YamlWorkloadCheck.Run();
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(YamlConformanceComparison).Assembly).Run(args);
