using BenchmarkDotNet.Running;
using OfficeIMO.Adf.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    return AdfEvidenceRunner.RunProbe(args[1..]);
}

if (args.Length > 0 && string.Equals(args[0], "evidence", StringComparison.OrdinalIgnoreCase)) {
    return AdfEvidenceRunner.Run(args[1..]);
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
return 0;
