using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
