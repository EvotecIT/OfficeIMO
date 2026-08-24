using BenchmarkDotNet.Running;
using OfficeIMO.AsciiDoc.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = AsciiDocEvidenceRunner.RunProbe(args[1..]);
    return;
}

if (args.Any(argument => string.Equals(argument, "--verify-budgets", StringComparison.OrdinalIgnoreCase))) {
    Environment.ExitCode = AsciiDocEvidenceRunner.RunEvidence(args, verifyBudgets: true);
    return;
}

if (args.Any(argument => string.Equals(argument, "--evidence", StringComparison.OrdinalIgnoreCase))) {
    Environment.ExitCode = AsciiDocEvidenceRunner.RunEvidence(args, verifyBudgets: false);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
    AsciiDocBenchmarkValidation.ValidateAll();
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
