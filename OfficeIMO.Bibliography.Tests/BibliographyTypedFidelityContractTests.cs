using OfficeIMO;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyTypedFidelityContractTests {
    [Theory]
    [InlineData(BibliographyConversionAction.Mapped, BibliographyDiagnosticSeverity.Information, OfficeConversionLossKind.None)]
    [InlineData(BibliographyConversionAction.PreservedExtension, BibliographyDiagnosticSeverity.Information, OfficeConversionLossKind.None)]
    [InlineData(BibliographyConversionAction.Approximated, BibliographyDiagnosticSeverity.Warning, OfficeConversionLossKind.Approximation)]
    [InlineData(BibliographyConversionAction.Omitted, BibliographyDiagnosticSeverity.Warning, OfficeConversionLossKind.Omission)]
    [InlineData(BibliographyConversionAction.Mapped, BibliographyDiagnosticSeverity.Error, OfficeConversionLossKind.Failure)]
    public void ReportProjectsExactLossCategoryIntoTheCommonContract(
        BibliographyConversionAction action,
        BibliographyDiagnosticSeverity severity,
        OfficeConversionLossKind expected) {
        var report = new BibliographyConversionReport();
        report.Add(new BibliographyConversionDiagnostic(
            "BIBTEST001", severity, "Evidence", action, "cite-key", "title"));

        OfficeConversionReport common = report;
        OfficeDiagnostic diagnostic = Assert.Single(common.Diagnostics);

        Assert.Equal(expected, diagnostic.LossKind);
        Assert.Equal("OfficeIMO.Bibliography", diagnostic.Format);
        Assert.Equal("cite-key:title", diagnostic.Location);
        Assert.Equal(expected != OfficeConversionLossKind.None, common.HasLoss);
        Assert.Equal(expected, Assert.Single(OfficeConversionFidelityDiagnostics.Flatten(new[] { common })).LossKind);
    }


    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("BIBTEST002")]
    public void BlankLegacyCodeRemainsReadableThroughTypedAndComposedReports(string code) {
        var report = new BibliographyConversionReport();
        report.Add(new BibliographyConversionDiagnostic(
            code, BibliographyDiagnosticSeverity.Warning, "Source field omitted.", BibliographyConversionAction.Omitted));

        Assert.Equal(code, Assert.Single(report.NativeDiagnostics).Code);
        OfficeDiagnostic typed = Assert.Single(report.Diagnostics);
        Assert.Equal(string.IsNullOrWhiteSpace(code) ? "BIBLIOGRAPHY_DIAGNOSTIC" : code, typed.Code);
        Assert.Equal(OfficeConversionLossKind.Omission, typed.LossKind);
        Assert.Equal(typed.Code, Assert.Single(OfficeConversionFidelityDiagnostics.Flatten(new OfficeConversionReport[] { report })).Code);
        Assert.Throws<BibliographyConversionLossException>(() => report.ThrowIfLoss());
    }
}
