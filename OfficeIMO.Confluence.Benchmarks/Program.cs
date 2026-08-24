using BenchmarkDotNet.Running;
using OfficeIMO.Confluence.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "evidence", StringComparison.OrdinalIgnoreCase)) {
    return ConfluenceEvidenceRunner.RunEvidence(args.Skip(1).ToArray());
}

if (args.Length > 0 && string.Equals(args[0], "evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    return ConfluenceEvidenceRunner.RunProbe(args.Skip(1).ToArray());
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
return 0;
