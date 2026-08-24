using BenchmarkDotNet.Running;
using OfficeIMO.Html.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--layout-evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlLayoutEvidenceRunner.RunProbe(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--layout-evidence", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlLayoutEvidenceRunner.RunEvidence(args.Skip(1).ToArray(), verifyBudgets: false);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--layout-verify-budgets", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlLayoutEvidenceRunner.RunEvidence(args.Skip(1).ToArray(), verifyBudgets: true);
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
