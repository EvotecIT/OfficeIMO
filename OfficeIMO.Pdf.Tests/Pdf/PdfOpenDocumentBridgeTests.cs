using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.OpenDocument;
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
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Load(source);

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
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Load(source);

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
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Load(source);

        PdfOdsConversionResult result = pdf.ToOdsDocumentResult();

        Assert.NotEmpty(result.Value.Sheets);
        Assert.True(result.Report.PdfReport.HasOmittedPageContent);
        Assert.True(result.HasLoss);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Equal("PK", System.Text.Encoding.ASCII.GetString(result.Value.ToBytes(), 0, 2));
    }

    [Fact]
    public void PdfToOdp_DefaultsToEditableContentAndProducesValidOdfPackage() {
        byte[] source = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .H1("First page")
            .Table(new[] {
                new[] { "Metric", "Value", "Status" },
                new[] { "Ready", "Yes", "Current" },
                new[] { "Loss", "Reported", "Current" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 100, 100, 120 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .PageBreak()
            .H1("Second page")
            .ToBytes();
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Load(source);

        PdfOdpConversionResult result = pdf.ToOdpPresentationResult();

        Assert.Equal(PdfPowerPointImportMode.EditableContent, result.Report.PdfReport.Mode);
        Assert.Equal(2, result.Value.Slides.Count);
        Assert.Equal(2, result.Report.PdfReport.EditablePages.Count);
        Assert.All(result.Report.PdfReport.EditablePages, page => Assert.True(page.TextBoxCount >= 1));
        byte[] artifact = result.Value.ToBytes();
        Assert.Equal("PK", System.Text.Encoding.ASCII.GetString(artifact, 0, 2));

        OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(artifact));
        Assert.Contains(reopened.Slides.SelectMany(slide => slide.Shapes).OfType<OdpTextBox>(),
            textBox => textBox.Paragraphs.Any(paragraph => paragraph.Text.Contains("First page", StringComparison.Ordinal)));
        OdpTable table = Assert.Single(reopened.Slides.SelectMany(slide => slide.Shapes).OfType<OdpTable>());
        Assert.Equal("Ready", table.Cell(1, 0).Text);
        Assert.Equal("Yes", table.Cell(1, 1).Text);
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
