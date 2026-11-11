using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentPdfConversionContracts {
    [Fact]
    public void OdtFacadePreservesProjectionLossAndProducesReadablePdf() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Direct ODT PDF");
        source.AddTrackedParagraphInsertion("Tracked source", "Reviewer");

        using PdfDocument pdf = source.ToPdf(null, out OfficeConversionReport report);
        byte[] bytes = pdf.ToBytes(new PdfSaveOptions { AllowLoss = true });

        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        Assert.Contains("Direct ODT PDF", PdfDocument.Load(new System.IO.MemoryStream(bytes)).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(report.Diagnostics, diagnostic =>
            diagnostic.Location == "source-tracked-changes" && diagnostic.LossKind != OfficeConversionLossKind.None);
        Assert.True(report.HasLoss);
        OdfConversionLossException loss = Assert.Throws<OdfConversionLossException>(() => source.ToPdf(new OdtPdfExportOptions {
            ConversionOptions = new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported }
        }, out _));
        Assert.Contains(loss.Report.Mappings, mapping =>
            mapping.Feature == "source-tracked-changes" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OfficeConversionException>(() => source.ToPdf(new OdtPdfExportOptions { FailOnLoss = true }, out _));
    }

    [Fact]
    public void OdsFacadeUsesExcelPdfEngineAndExposesInformationEvidence() {
        OdsDocument source = OdsDocument.Create();
        source.AddSheet("Revenue").Cell(0, 0).SetString("Quarter total");

        using PdfDocument pdf = source.ToPdf(null, out OfficeConversionReport report);
        byte[] bytes = pdf.ToBytes(new PdfSaveOptions { AllowLoss = true });
        string text = PdfDocument.Load(new System.IO.MemoryStream(bytes)).ExtractText();

        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        Assert.Contains("Quarter total", text, StringComparison.Ordinal);
        Assert.Single(PdfDocument.Load(new System.IO.MemoryStream(bytes)).Pages);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Failure);
    }

    [Fact]
    public void OdpFacadeUsesPowerPointPdfEngineAndKeepsAnimationLoss() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide("Overview");
        slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2), "Direct ODP PDF");
        OdpRectangle animated = slide.AddRectangle(OdfRect.FromCentimeters(1, 4, 2, 2));
        slide.AddFadeInAnimation(animated, TimeSpan.FromSeconds(1));

        using PdfDocument pdf = source.ToPdf(null, out OfficeConversionReport report);
        byte[] bytes = pdf.ToBytes(new PdfSaveOptions { AllowLoss = true });
        string text = PdfDocument.Load(new System.IO.MemoryStream(bytes)).ExtractText();

        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        Assert.Contains("Direct ODP PDF", text, StringComparison.Ordinal);
        Assert.Contains(report.Diagnostics, diagnostic =>
            diagnostic.Location == "source-presentation-animations" && diagnostic.LossKind != OfficeConversionLossKind.None);
        Assert.True(report.HasLoss);
    }

    [Fact]
    public void OdtPdfRepeatsRunningHeadersAndKeepsTablesImagesAndPageBreaks() {
        OdtDocument source = OdtDocument.Create();
        source.PageLayout.Header.AddParagraph("Quarterly header");
        source.PageLayout.Footer.AddParagraph("Confidential footer");
        source.AddHeading("Quarterly review", 1);
        OdtTable table = source.AddTable(2, 2, "Totals");
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Region";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Revenue";
        table.Rows[1].Cells[0].Paragraphs[0].Text = "North";
        table.Rows[1].Cells[1].Paragraphs[0].Text = "1200";
        source.AddParagraph().AddImage(OnePixelPng, "logo.png", OdfLength.Centimeters(2), OdfLength.Centimeters(2));
        source.AddPageBreak();
        source.AddParagraph("Closing page");

        using PdfDocument pdf = source.ToPdf(null, out OfficeConversionReport report);
        // Running headers and footers are tagged as pagination artifacts.
        PdfDocument reopened = PdfDocument.Load(new System.IO.MemoryStream(pdf.ToBytes()),
            new PdfLoadOptions { IncludeArtifactText = true });

        Assert.Equal(2, reopened.Pages.Count);
        Assert.Equal("Quarterly header Quarterly review Region Revenue North 1200 Confidential footer",
            NormalizedText(reopened.Pages[0].ExtractText()));
        Assert.Equal("Quarterly header Closing page Confidential footer", NormalizedText(reopened.Pages[1].ExtractText()));
        Assert.Single(reopened.ExtractImages());
        Assert.False(report.HasLoss);
    }

    [Fact]
    public void OdsPdfPrintsOnlyTheSheetPrintRange() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Revenue");
        for (int row = 0; row < 60; row++) {
            sheet.Cell(row, 0).SetString("Row " + (row + 1));
            sheet.Cell(row, 1).SetNumber(row * 10);
            sheet.Cell(row, 5).SetString("Outside " + (row + 1));
        }
        sheet.PrintRanges = "Revenue.A1:Revenue.B3";

        using PdfDocument pdf = source.ToPdf(null, out OfficeConversionReport report);
        PdfDocument reopened = PdfDocument.Load(new System.IO.MemoryStream(pdf.ToBytes()));

        Assert.Equal("Row 1 0 Row 2 10 Row 3 20", NormalizedText(Assert.Single(reopened.Pages).ExtractText()));
        Assert.False(report.HasLoss);
    }

    [Fact]
    public void OdpPdfRendersOneSlideSizedPagePerSlideWithTablesAndImages() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide first = source.AddSlide("Overview");
        first.AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 2), "Quarterly overview");
        OdpTable table = first.AddTable(OdfRect.FromCentimeters(1, 4, 10, 3), 2, 2);
        table.Cell(0, 0).Text = "Region";
        table.Cell(0, 1).Text = "Revenue";
        table.Cell(1, 0).Text = "North";
        table.Cell(1, 1).Text = "1200";
        first.AddImage(OnePixelPng, "logo.png", OdfRect.FromCentimeters(14, 1, 3, 3));
        source.AddSlide("Details").AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 2), "Second slide details");

        using PdfDocument pdf = source.ToPdf(null, out OfficeConversionReport report);
        PdfDocument reopened = PdfDocument.Load(new System.IO.MemoryStream(pdf.ToBytes()));

        Assert.Equal(2, reopened.Pages.Count);
        Assert.Equal(source.PageWidth.ToPoints(), reopened.Pages[1].GetPageSize().Width, 0);
        Assert.Equal(source.PageHeight.ToPoints(), reopened.Pages[1].GetPageSize().Height, 0);
        Assert.Equal("Quarterly overview Region Revenue North 1200", NormalizedText(reopened.Pages[0].ExtractText()));
        Assert.Equal("Second slide details", NormalizedText(reopened.Pages[1].ExtractText()));
        Assert.Single(reopened.ExtractImages());
        Assert.False(report.HasLoss);
    }

    private static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC");

    private static string NormalizedText(string text) =>
        System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ").Trim();

    [Fact]
    public async Task DirectFacadeSupportsStreamAndAsyncSaveContracts() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Stream contract");
        source.AddTrackedParagraphInsertion("Tracked source", "Reviewer");
        using var stream = new MemoryStream();

        OfficeConversionReport save = await source.SaveAsPdfAsync(stream);

        Assert.True(stream.Length > 0);
        Assert.True(stream.CanWrite);
        Assert.True(save.HasLoss);
        Assert.Contains(save.Diagnostics, diagnostic => diagnostic.Location == "source-tracked-changes");
    }
}
