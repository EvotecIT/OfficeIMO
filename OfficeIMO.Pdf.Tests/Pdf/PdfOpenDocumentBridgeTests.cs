using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfOpenDocumentBridgeTests {
    [Fact]
    public void PdfToOdt_ReconstructsSemanticTextAndExposesBothStageReports() {
        byte[] source = PdfCore.PdfDocument.Create()
            .H1("Quarterly summary")
            .Paragraph(paragraph => paragraph.Text("Revenue increased."))
            .ToBytes();
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Open(source);

        PdfOdtConversionResult result = pdf.ToOdtDocumentResult();
        byte[] odt = result.Value.ToBytes();

        Assert.Contains(result.Value.Paragraphs, paragraph =>
            paragraph.Text.Contains("Revenue increased.", StringComparison.Ordinal));
        Assert.NotNull(result.Report.PdfReport);
        Assert.NotNull(result.Report.OpenDocumentReport);
        Assert.True(odt.Length > 100);
        Assert.Equal("PK", System.Text.Encoding.ASCII.GetString(odt, 0, 2));
    }

    [Fact]
    public void PdfToOdt_ReportsLostFormInteractivityThroughTheCombinedLossGate() {
        byte[] source = PdfCore.PdfDocument.Create()
            .TextField("Approval", width: 120, value: "Ready")
            .ToBytes();
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Open(source);

        PdfOdtConversionResult result = pdf.ToOdtDocumentResult();

        Assert.True(result.HasLoss);
        Assert.Contains(result.Report.PdfReport.Warnings, warning =>
            warning.Code == "PdfFormWidgetPlaceholder" &&
            warning.Severity == PdfCore.PdfConversionWarningSeverity.Warning);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
    }

    [Fact]
    public void PdfToOds_ReconstructsDetectedTablesAndReportsOmittedPageContentAsLoss() {
        byte[] source = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Invoice summary"))
            .KeyValueTable(new[] {
                PdfCore.PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfCore.PdfKeyValueRow.Text("Customer", "Evotec")
            })
            .ToBytes();
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Open(source);

        PdfOdsConversionResult result = pdf.ToOdsDocumentResult();

        Assert.NotEmpty(result.Value.Sheets);
        Assert.True(result.Report.PdfReport.HasOmittedPageContent);
        Assert.True(result.HasLoss);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Equal("PK", System.Text.Encoding.ASCII.GetString(result.Value.ToBytes(), 0, 2));
    }

    [Fact]
    public void PdfToOdp_DefaultsToOneVisualSlidePerPageAndProducesValidOdfPackage() {
        byte[] source = PdfCore.PdfDocument.Create()
            .H1("First page")
            .PageBreak()
            .H1("Second page")
            .ToBytes();
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Open(source);

        PdfOdpConversionResult result = pdf.ToOdpPresentationResult();

        Assert.Equal(PdfPowerPointImportMode.VisualPages, result.Report.PdfReport.Mode);
        Assert.Equal(2, result.Value.Slides.Count);
        Assert.Equal(2, result.Report.PdfReport.VisualPages.Count);
        Assert.Equal("PK", System.Text.Encoding.ASCII.GetString(result.Value.ToBytes(), 0, 2));
    }

    [Theory]
    [InlineData(typeof(OdtPdfConversionExtensions), "SaveAsOdt", "ToOdtDocument", "ToOdtDocumentResult")]
    [InlineData(typeof(OdsPdfConversionExtensions), "SaveAsOds", "ToOdsDocument", "ToOdsDocumentResult")]
    [InlineData(typeof(OdpPdfConversionExtensions), "SaveAsOdp", "ToOdpPresentation", "ToOdpPresentationResult")]
    public void ReverseOpenDocumentAdaptersUseTheSameFacadeOnOpenedAndLogicalPdfDocuments(
        Type converterType,
        string saveName,
        string importName,
        string resultName) {
        System.Reflection.MethodInfo[] methods = converterType
            .GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);

        Assert.Equal(4, methods.Count(method => method.Name == saveName));
        Assert.Equal(4, methods.Count(method => method.Name == saveName + "Async"));
        Assert.Equal(2, methods.Count(method => method.Name == importName));
        Assert.Equal(2, methods.Count(method => method.Name == resultName));
    }
}
