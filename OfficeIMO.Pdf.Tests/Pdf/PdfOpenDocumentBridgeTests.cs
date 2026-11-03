using OfficeIMO.Pdf.PowerPoint;
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

        PdfOdtConversionResult result = pdf.ConvertToOdtWithEvidence(null, default);
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
            .TextField("Approval", width: global::OfficeIMO.OfficeLength.Points(120), value: "Ready")
            .ToBytes();
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Load(source);

        PdfOdtConversionResult result = pdf.ConvertToOdtWithEvidence(null, default);

        Assert.True(result.HasLoss);
        Assert.Contains(result.Report.PdfReport.Warnings, warning =>
            warning.Code == "pdf-form-widget-placeholder" &&
            warning.Severity == PdfCore.PdfConversionWarningSeverity.Warning);
        Assert.ThrowsAny<OfficeConversionException>(() => result.ThrowIfLoss());
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

        PdfOdsConversionResult result = pdf.ConvertToOdsWithEvidence(null, default);

        Assert.NotEmpty(result.Value.Sheets);
        Assert.True(result.Report.PdfReport.HasOmittedPageContent);
        Assert.True(result.HasLoss);
        Assert.ThrowsAny<OfficeConversionException>(() => result.ThrowIfLoss());
        Assert.Equal("PK", System.Text.Encoding.ASCII.GetString(result.Value.ToBytes(), 0, 2));
    }

    [Fact]
    public void PdfToOdp_DefaultsToEditableContentAndProducesValidOdfPackage() {
        byte[] source = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(360),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .H1("First page")
            .Table(new[] {
                new[] { "Metric", "Value", "Status" },
                new[] { "Ready", "Yes", "Current" },
                new[] { "Loss", "Reported", "Current" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(100), global::OfficeIMO.OfficeLength.Points(100), global::OfficeIMO.OfficeLength.Points(120) },
                HeaderRowCount = 1,
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(4)
            })
            .PageBreak()
            .H1("Second page")
            .ToBytes();
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Load(source);

        PdfOdpConversionResult result = pdf.ConvertToOdpWithEvidence(null, default);

        Assert.Equal(PowerPointPdfImportMode.EditableContent, result.Report.PdfReport.Mode);
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

}
