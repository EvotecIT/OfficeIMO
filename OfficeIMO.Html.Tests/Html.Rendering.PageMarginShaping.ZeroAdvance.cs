using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("min-content")]
    [InlineData("max-content")]
    [InlineData("fit-content")]
    public void HtmlPageMarginShaping_ZeroAdvanceKeepsItsIntrinsicBoxAndRetainedGlyph(string keyword) {
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(
            PageMarginZeroAdvanceDocument(keyword), PageMarginShapingOptions(cff: true));
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        HtmlRenderText text = Assert.Single(PageMarginShapedText(box));

        Assert.NotNull(text.PreparedShaping);
        Assert.Equal("\u0301", text.PreparedShaping!.Text);
        Assert.NotEmpty(text.PreparedShaping.Result.Glyphs);
        Assert.Equal(0D, text.PreparedShaping.Measure(text.Font.Size));
        Assert.Equal(0D, text.TextAdvanceWidth);
        Assert.Equal(14D, box.Width, 6);
        Assert.DoesNotContain(EnumerateRenderVisuals(box.Visuals), visual => visual is HtmlRenderClipGroup);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("min-content")]
    [InlineData("max-content")]
    [InlineData("fit-content")]
    public void HtmlPageMarginShaping_SubpixelAdvanceDoesNotAcquireAOnePixelFloor(string keyword) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(PageMarginShapingSource("top-center",
            "A\u0301", "width:" + keyword + ";font-size:0.5px;padding:5px;border:2px solid red", string.Empty));
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, PageMarginShapingOptions(cff: true));
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        HtmlRenderText text = Assert.Single(PageMarginShapedText(box));

        Assert.NotNull(text.PreparedShaping);
        double advance = text.PreparedShaping!.Measure(text.Font.Size);
        Assert.InRange(advance, double.Epsilon, 0.999D);
        Assert.Equal(14D + advance, box.Width, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginShaping_ZeroAdvanceStillHonorsAnAuthoredZeroHeight() {
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(
            PageMarginZeroAdvanceDocument("max-content", ";height:0"), PageMarginShapingOptions(cff: true));
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        HtmlRenderClipGroup clip = Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderClipGroup>());

        Assert.Equal(0D, clip.ClipHeight);
        Assert.True(clip.ClipVertical);
        Assert.False(clip.ClipHorizontal);
        Assert.Contains(rendered.Diagnostics, item => item.Source == "@page @top-center"
            && item.Detail == "margin-height" && item.LossKind == OfficeConversionLossKind.Omission);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlPageMarginShaping_ZeroAdvanceExportsTheSameSavedPdfAndImageScene() {
        HtmlConversionDocument document = PageMarginZeroAdvanceDocument("max-content");
        HtmlToPdfOptions options = PageMarginShapingOptions(cff: true);
        PdfCore.PdfDocumentConversionResult pdf = document.ConvertToPdfWithEvidence(options);
        pdf.RequireNoLoss();
        byte[] bytes = pdf.ToBytes();
        pdf.RequireNoLoss();
        PdfCore.PdfReadDocument saved = PdfCore.PdfReadDocument.Open(bytes,
            new PdfCore.PdfLoadOptions { IncludeArtifactText = true });
        Assert.Equal("\u0301", saved.ExtractText().Trim());

        OfficeImageExportResult image = document.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult replay = Assert.Single(PdfCore.PdfImageExportExtensions.ExportImages(pdf,
            OfficeImageExportFormat.Png, new PdfCore.PdfImageExportOptions { Scale = 96D / 72D }));
        Assert.Equal(replay.Bytes, image.Bytes);
        Assert.DoesNotContain(image.Diagnostics, item => item.LossKind != OfficeConversionLossKind.None);
        Assert.DoesNotContain(replay.Diagnostics, item => item.LossKind != OfficeConversionLossKind.None);
    }

    private static HtmlConversionDocument PageMarginZeroAdvanceDocument(string keyword, string extra = "") =>
        HtmlConversionDocument.Parse(PageMarginShapingSource("top-center", "\u0301",
            "width:" + keyword + ";padding:5px;border:2px solid red" + extra, string.Empty));
}
