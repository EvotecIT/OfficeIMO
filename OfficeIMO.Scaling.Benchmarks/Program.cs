using BenchmarkDotNet.Running;
using OfficeIMO.Scaling.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--scan", StringComparison.OrdinalIgnoreCase)) {
    return ScalingScan.Run(args.Skip(1).ToArray());
}

BenchmarkSwitcher.FromAssembly(typeof(FormatScalingBenchmarks).Assembly).Run(args);
return 0;
