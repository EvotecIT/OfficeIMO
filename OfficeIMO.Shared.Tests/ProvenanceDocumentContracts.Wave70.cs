using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void OverriddenCssDataUrlIsNotReportedAsActiveProvenance() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<style>.hero{{background-image:url('{dataUri}')}}.hero{{background-image:url(clean.png)}}</style><div class='hero'></div>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void HtmlPreflightDoesNotChargeIgnoredNestedForms() {
        const string html = "<form><form><form><form><form>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html,
            new OfficeProvenanceOptions { MaxContainerEntries = 4 });

        Assert.Empty(report.Evidence);
    }

    [Theory]
    [InlineData("xl%2Fworkbook.bin")]
    [InlineData("xl%5cworkbook.bin")]
    public void ExcelXlsbRejectsPercentEncodedPackageSeparators(string target) {
        byte[] package = CreateWave33XlsbProvenancePackage(signed: false, officeDocumentTarget: target);

        Assert.ThrowsAny<Exception>(() => ExcelDocument.RemoveProvenance(package, "workbook.xlsb"));
    }
}
