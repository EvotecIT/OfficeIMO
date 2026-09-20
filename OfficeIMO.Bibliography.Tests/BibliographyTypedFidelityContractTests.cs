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

        IOfficeConversionReport common = report;
        OfficeConversionFidelityDiagnostic diagnostic = Assert.Single(common.FidelityDiagnostics);

        Assert.Equal(expected, diagnostic.LossKind);
        Assert.Equal("OfficeIMO.Bibliography", diagnostic.Source);
        Assert.Equal("cite-key:title", diagnostic.Location);
        Assert.Equal(expected != OfficeConversionLossKind.None, common.HasLoss);
        Assert.Equal(expected, Assert.Single(OfficeConversionFidelityDiagnostics.Flatten(new[] { common })).LossKind);
    }

    [Fact]
    public void StrictAcceptanceRejectsTypedBibliographyLoss() {
        var report = new BibliographyConversionReport();
        report.Add(new BibliographyConversionDiagnostic(
            "BIBTEST002",
            BibliographyDiagnosticSeverity.Warning,
            "Source field omitted.",
            BibliographyConversionAction.Omitted));

        Assert.Throws<BibliographyConversionLossException>(() => report.RequireNoLoss());
    }
}
