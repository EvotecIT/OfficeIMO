using BenchmarkDotNet.Running;
using OfficeIMO.Html.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--qualification-baseline", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlQualificationBaselineRunner.Run(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--layout-fingerprint", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlLayoutFingerprintRunner.Run(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--layout-evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlLayoutEvidenceRunner.RunProbe(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--layout-evidence", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlLayoutEvidenceRunner.RunEvidence(args.Skip(1).ToArray(), verifyBudgets: false);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--provider-evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlProviderEvidenceRunner.RunProbe(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--provider-evidence", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlProviderEvidenceRunner.Run(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--owned-document-evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlOwnedDocumentEvidenceRunner.RunProbe(args.Skip(1).ToArray());
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--owned-document-evidence", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlOwnedDocumentEvidenceRunner.Run(args.Skip(1).ToArray(), verifyBudgets: false);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--owned-document-verify-budgets", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlOwnedDocumentEvidenceRunner.Run(args.Skip(1).ToArray(), verifyBudgets: true);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--layout-verify-budgets", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = HtmlLayoutEvidenceRunner.RunEvidence(args.Skip(1).ToArray(), verifyBudgets: true);
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
