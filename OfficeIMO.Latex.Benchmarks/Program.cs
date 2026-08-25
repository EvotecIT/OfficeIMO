using BenchmarkDotNet.Running;
using OfficeIMO.Latex.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = LatexEvidenceRunner.RunProbe(args[1..]);
    return;
}

if (args.Any(argument => string.Equals(argument, "--verify-budgets", StringComparison.OrdinalIgnoreCase))) {
    Environment.ExitCode = LatexEvidenceRunner.RunEvidence(args, verifyBudgets: true);
    return;
}

if (args.Any(argument => string.Equals(argument, "--evidence", StringComparison.OrdinalIgnoreCase))) {
    Environment.ExitCode = LatexEvidenceRunner.RunEvidence(args, verifyBudgets: false);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
    LatexBenchmarkValidation.ValidateAll();
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
