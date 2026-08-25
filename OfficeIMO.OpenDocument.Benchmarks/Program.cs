using BenchmarkDotNet.Running;
using OfficeIMO.OpenDocument.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--probe", StringComparison.OrdinalIgnoreCase)) {
    return OpenDocumentEvidenceRunner.RunProbe(args.Skip(1).ToArray());
}

if (args.Any(argument => string.Equals(argument, "--evidence", StringComparison.OrdinalIgnoreCase))) {
    return OpenDocumentEvidenceRunner.RunEvidence(args);
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
return 0;
