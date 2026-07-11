using BenchmarkDotNet.Running;
using OfficeIMO.Rtf.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--probe", StringComparison.OrdinalIgnoreCase)) {
    return RtfBenchmarkBudgetRunner.RunProbe(args.Skip(1).ToArray());
}

if (args.Length > 0 && string.Equals(args[0], "--verify-budgets", StringComparison.OrdinalIgnoreCase)) {
    return RtfBenchmarkBudgetRunner.Verify(args.Skip(1).ToArray());
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
return 0;
