using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void ExcelXlsbRejectsAbsoluteInternalWorkbookTargets() {
        byte[] package = CreateWave33XlsbProvenancePackage(
            signed: false,
            officeDocumentTarget: "http://package/xl/workbook.bin");

        Exception? exception = Record.Exception(() =>
            ExcelDocument.RemoveProvenance(package, "workbook.xlsb"));

        Assert.NotNull(exception);
        Assert.True(
            exception is InvalidDataException or ArgumentException,
            $"Unexpected package-rejection exception: {exception}");
    }

    [Fact]
    public void HtmlPreflightDoesNotChargeIgnoredDuplicateDocumentElements() {
        const string html = "<html><head></head><body><body></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html,
            new OfficeProvenanceOptions { MaxContainerEntries = 3 });

        Assert.Empty(report.Evidence);
        Assert.Equal(OfficeProvenanceAssetFormat.Html, report.Format);
    }
}
