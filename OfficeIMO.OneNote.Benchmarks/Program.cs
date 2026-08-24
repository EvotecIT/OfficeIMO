using BenchmarkDotNet.Running;
using OfficeIMO.OneNote.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--probe", StringComparison.OrdinalIgnoreCase)) {
    return OneNoteEvidenceRunner.RunProbe(args.Skip(1).ToArray());
}

if (args.Any(argument => string.Equals(argument, "--verify-budgets", StringComparison.OrdinalIgnoreCase))) {
    return OneNoteEvidenceRunner.RunEvidence(args, verifyBudgets: true);
}

if (args.Any(argument => string.Equals(argument, "--evidence", StringComparison.OrdinalIgnoreCase))) {
    return OneNoteEvidenceRunner.RunEvidence(args, verifyBudgets: false);
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
return 0;
