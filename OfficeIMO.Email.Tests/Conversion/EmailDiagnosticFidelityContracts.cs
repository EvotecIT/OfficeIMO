using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailDiagnosticFidelityContracts {
    [Theory]
    [InlineData(EmailDiagnosticSeverity.Information)]
    [InlineData(EmailDiagnosticSeverity.Warning)]
    public void UnassessedDiagnosticsRemainUnassessedAndFailStrictAcceptance(EmailDiagnosticSeverity severity) {
        var simple = new EmailDiagnostic("source-rendering", "Fidelity was not assessed.", severity,
            "body", OfficeConversionLossKind.Unassessed);
        var actionable = CreateActionable(severity, EmailDiagnosticDisposition.Observed,
            OfficeConversionLossKind.Unassessed);
        var report = new EmailConversionReport(EmailFileFormat.Eml, EmailFileFormat.Eml,
            new[] { simple, actionable });

        Assert.Equal(OfficeConversionLossKind.Unassessed, simple.LossKind);
        Assert.Equal(OfficeConversionLossKind.Unassessed, actionable.LossKind);
        Assert.All(report.FidelityDiagnostics, diagnostic =>
            Assert.Equal(OfficeConversionLossKind.Unassessed, diagnostic.LossKind));
        Assert.True(report.HasLoss);
        Assert.True(report.CanWrite);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());
    }

    [Theory]
    [InlineData(EmailDiagnosticSeverity.Error, EmailDiagnosticDisposition.Observed)]
    [InlineData(EmailDiagnosticSeverity.Information, EmailDiagnosticDisposition.Stopped)]
    public void ErrorOrStoppedProcessingStillRequiresFailureClassification(
        EmailDiagnosticSeverity severity, EmailDiagnosticDisposition disposition) {
        EmailDiagnostic diagnostic = CreateActionable(severity, disposition, OfficeConversionLossKind.Unassessed);
        var report = new EmailConversionReport(EmailFileFormat.Eml, EmailFileFormat.Eml, new[] { diagnostic });

        Assert.Equal(OfficeConversionLossKind.Failure, diagnostic.LossKind);
        Assert.False(report.CanWrite);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(999)]
    public void BothConstructorsRejectUndefinedFidelityCategories(int value) {
        Assert.Throws<ArgumentOutOfRangeException>(() => new EmailDiagnostic("source-rendering", "Invalid category.",
            EmailDiagnosticSeverity.Information, "body", (OfficeConversionLossKind)value));
        Assert.Throws<ArgumentOutOfRangeException>(() => CreateActionable(EmailDiagnosticSeverity.Information,
            EmailDiagnosticDisposition.Observed, (OfficeConversionLossKind)value));
    }

    private static EmailDiagnostic CreateActionable(EmailDiagnosticSeverity severity,
        EmailDiagnosticDisposition disposition, OfficeConversionLossKind lossKind) =>
        new("source-rendering", "Fidelity was not assessed.", severity, "body", "conversion", null, null,
            null, null, disposition, EmailDataLossRisk.None, "Assess fidelity before strict acceptance.",
            isRetryable: false, lossKind: lossKind);
}
