using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfConversionReportTests {
    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    public void BlankLegacyCodeProjectsAStableTypedIdentifier(string code) {
        var report = new RtfConversionReport();
        report.Add(RtfConversionSeverity.Warning, code, "Content omitted.", RtfConversionAction.Omitted, "body/0");

        Assert.Equal(code, Assert.Single(report.Diagnostics).Code);
        OfficeConversionFidelityDiagnostic typed = Assert.Single(
            OfficeConversionFidelityDiagnostics.Flatten(new[] { report }));
        Assert.Equal("RTF_DIAGNOSTIC_UNSPECIFIED", typed.Code);
        Assert.Equal(OfficeConversionLossKind.Omission, typed.LossKind);
        Assert.Equal("body/0", typed.Location);
        Assert.True(report.HasLoss);
        Assert.Throws<RtfConversionLossException>(report.RequireNoLoss);
    }

    [Fact]
    public void Preserved_And_Substituted_Actions_Do_Not_Fail_Strict_Mode() {
        var report = new RtfConversionReport();
        report.Add(RtfConversionSeverity.Information, "Preserved", "Preserved.", RtfConversionAction.Preserved);
        report.Add(RtfConversionSeverity.Warning, "Substituted", "Substituted.", RtfConversionAction.Substituted);

        report.RequireNoLoss();

        Assert.False(report.HasLoss);
        Assert.Equal(2, report.Diagnostics.Count);
        Assert.All(report.FidelityDiagnostics, diagnostic => Assert.Equal(OfficeConversionLossKind.None, diagnostic.LossKind));
    }

    [Theory]
    [InlineData(RtfConversionAction.Flattened, OfficeConversionLossKind.Approximation)]
    [InlineData(RtfConversionAction.Omitted, OfficeConversionLossKind.Omission)]
    [InlineData(RtfConversionAction.Blocked, OfficeConversionLossKind.Failure)]
    public void Loss_Actions_Fail_Strict_Mode(
        RtfConversionAction action,
        OfficeConversionLossKind expectedLossKind) {
        var report = new RtfConversionReport();
        report.Add(RtfConversionSeverity.Warning, "Loss", "Loss occurred.", action, "Body/0", "feature", 2, "detail");

        RtfConversionLossException exception = Assert.Throws<RtfConversionLossException>(() => report.RequireNoLoss());

        Assert.Same(report, exception.Report);
        RtfConversionDiagnostic diagnostic = Assert.Single(report.Diagnostics);
        Assert.Equal("Body/0", diagnostic.SourcePath);
        Assert.Equal("feature", diagnostic.Feature);
        Assert.Equal(2, diagnostic.Count);
        Assert.Equal("detail", diagnostic.Detail);
        Assert.Equal(expectedLossKind, Assert.Single(report.FidelityDiagnostics).LossKind);
    }

    [Fact]
    public void Generic_Result_Returns_Value_After_Strict_Check() {
        var report = new RtfConversionReport();
        var result = new RtfConversionResult<string>("value", report);

        Assert.True(result.Succeeded);
        Assert.False(result.HasLoss);
        Assert.Equal("value", result.RequireValue());
        Assert.Equal("value", result.RequireNoLoss());
    }

    [Fact]
    public void Generic_Result_Distinguishes_Errors_From_Fidelity_Loss() {
        var report = new RtfConversionReport();
        report.Add(RtfConversionSeverity.Warning, "Flattened", "Flattened.", RtfConversionAction.Flattened);
        var result = new RtfConversionResult<string>("value", report);

        Assert.True(result.Succeeded);
        Assert.True(result.HasLoss);
        Assert.Equal("value", result.RequireValue());
        Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
    }

    [Theory]
    [InlineData("RTF012", RtfDiagnosticSeverity.Warning, RtfConversionAction.Flattened, OfficeConversionLossKind.Approximation)]
    [InlineData("RTF103", RtfDiagnosticSeverity.Warning, RtfConversionAction.Flattened, OfficeConversionLossKind.Approximation)]
    [InlineData("RTF999", RtfDiagnosticSeverity.Warning, RtfConversionAction.Flattened, OfficeConversionLossKind.Approximation)]
    [InlineData("RTF101", RtfDiagnosticSeverity.Warning, RtfConversionAction.Omitted, OfficeConversionLossKind.Omission)]
    [InlineData("RTF105", RtfDiagnosticSeverity.Warning, RtfConversionAction.Blocked, OfficeConversionLossKind.Failure)]
    [InlineData("RTF013", RtfDiagnosticSeverity.Error, RtfConversionAction.Blocked, OfficeConversionLossKind.Failure)]
    public void Read_Diagnostics_Preserve_Recovery_And_Omission_Categories(
        string code,
        RtfDiagnosticSeverity severity,
        RtfConversionAction expectedAction,
        OfficeConversionLossKind expectedLossKind) {
        var report = new RtfConversionReport();

        report.AddReadDiagnostics(new[] { new RtfDiagnostic(severity, code, "Read diagnostic.", 7) }, "fixture.rtf");

        Assert.Equal(expectedAction, Assert.Single(report.Diagnostics).Action);
        OfficeConversionFidelityDiagnostic diagnostic = Assert.Single(report.FidelityDiagnostics);
        Assert.Equal(expectedLossKind, diagnostic.LossKind);
        Assert.Equal("fixture.rtf", diagnostic.Location);
    }
}
