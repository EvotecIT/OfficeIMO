using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfFontCoverageTests {
    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    public void HtmlPdf_UncoveredUnicodePaintsReplacementWithDiagnosticAndKeepsSourceText(
        HtmlRenderIntentProfile profile) {
        const string html = "<p>A−B</p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var request = HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });

        HtmlPdfRenderRequestResult result = document.RenderToPdfResult(request);

        Assert.Contains("A−B", PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
        PdfCore.PdfConversionWarning warning = Assert.Single(result.Output.Report.Warnings,
            item => item.Code == HtmlPdfDiagnosticCodes.UnavailableTextGlyphReplaced);
        Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.Contains("U+2212", warning.Message, StringComparison.Ordinal);
        Assert.Equal(html, document.SourceHtml);
    }

    [Fact]
    public void HtmlPdf_UncoveredUnicodeStillFailsWhenLossIsNotAllowed() {
        var request = HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        });

        HtmlConversionException failure = Assert.Throws<HtmlConversionException>(() =>
            HtmlConversionDocument.Parse("<p>A−B</p>").RenderToPdfDocumentResult(request));

        HtmlDiagnostic diagnostic = Assert.Single(failure.Diagnostics, item => item.Code == "unsupported-text-glyph");
        Assert.Equal(OfficeConversionLossKind.Failure, diagnostic.LossKind);
        Assert.Contains("PdfCanvas", diagnostic.Detail, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void HtmlPdf_PortableRegisteredFamilySplitsMissingStyledGlyphToCoveredFace(bool sourceFont, bool requestedBold) {
        byte[] covering = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B', 0x2212);
        byte[] limited = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B');
        byte[] regular = requestedBold ? covering : limited;
        byte[] bold = requestedBold ? limited : covering;
        string paragraph = "<p style='font-family:Missing,Scoped;font-weight:" + (requestedBold ? "bold" : "normal") + "'>A−B</p>";
        string html = sourceFont
            ? "<style>@font-face{font-family:Scoped;src:url(data:font/ttf;base64,"
                + Convert.ToBase64String(regular) + ");font-weight:normal}@font-face{font-family:Scoped;src:url(data:font/ttf;base64,"
                + Convert.ToBase64String(bold) + ");font-weight:bold}</style>" + paragraph
            : paragraph;
        var options = new HtmlToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        if (!sourceFont) options.PdfOptions.RegisterNamedFontFamily(
            new PdfCore.PdfEmbeddedFontFamily("Scoped", regular, bold));

        byte[] pdf = (HtmlConversionDocument.Parse(html).ToPdf(options)).ToBytes();

        Assert.Contains("A−B", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(PdfCore.PdfDiagnostics.Analyze(pdf).EmbeddedFontCount >= 1);
    }

    [Fact]
    public void HtmlPdf_ColrV1GlyphPaintsItsLayersAndReportsFlattenedGradient() {
        byte[] font = ManagedTextShapingTestAssets.CreateColrV1Font('A');

        HtmlPdfRenderRequestResult result = RenderWithWebFont(font, "A");

        OfficeColor[] fills = DrawingTestTraversal.Elements(PdfCore.PdfDocument.Load(new System.IO.MemoryStream(result.ToBytes())).Render.Drawing(1))
            .OfType<OfficeDrawingShape>()
            .Select(shape => shape.Shape.FillColor ?? OfficeColor.Transparent)
            .ToArray();
        Assert.Contains(fills, color => color.R == 255 && color.G == 0 && color.B == 0);
        Assert.Contains(fills, color => color.R is > 120 and < 135 && color.G == 0 && color.B is > 120 and < 135);
        Assert.Contains(result.Output.Report.Warnings, item => item.Code == HtmlPdfDiagnosticCodes.ColorGlyphApproximated);
    }

    [Fact]
    public void HtmlPdf_BitmapOnlyColorGlyphReportsReplacementInsteadOfPaintingBlank() {
        const int emoji = 0x1F600;
        string text = char.ConvertFromUtf32(emoji);

        HtmlPdfRenderRequestResult result = RenderWithWebFont(ManagedTextShapingTestAssets.CreateBitmapOnlyColorFont(emoji), text);

        Assert.Contains(result.Output.Report.Warnings, item => item.Code == HtmlPdfDiagnosticCodes.UnavailableTextGlyphReplaced
            && item.Message.Contains("U+1F600", StringComparison.Ordinal));
        Assert.Contains(text, PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_ColorFlagLigatureIsNotSplitIntoItsRegionalIndicators() {
        string flag = char.ConvertFromUtf32(0x1F1F5) + char.ConvertFromUtf32(0x1F1F1);
        byte[] font = ManagedTextShapingTestAssets.CreateColorLigatureFont(0x1F1F5, 0x1F1F1);

        HtmlPdfRenderRequestResult result = RenderWithWebFont(font, flag);

        Assert.Contains(DrawingTestTraversal.Elements(PdfCore.PdfDocument.Load(new System.IO.MemoryStream(result.ToBytes())).Render.Drawing(1))
            .OfType<OfficeDrawingShape>(), shape => shape.Shape.FillColor?.B == 255);
        Assert.Contains(flag, PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_MarginBoxTextDoesNotTurnInstalledBoldIntoRegular() {
        if (!PdfCore.PdfEmbeddedFontFamily.TryFromSystem("Arial", out PdfCore.PdfEmbeddedFontFamily? arial) || arial?.Bold == null) return;
        const string html = "<style>@page{@bottom-center{content:'Page ' counter(page)}}body{font:12px Helvetica,sans-serif}</style><h1>Heading</h1><p>Body</p>";

        HtmlPdfRenderRequestResult result = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions()));

        result.ToBytes();
        string[] sources = result.Output.Report.Warnings.Where(item => item.Code == "font-embedding-cost")
            .Select(item => item.Message.Substring(item.Message.LastIndexOf(';'))).Distinct().ToArray();
        Assert.True(sources.Length >= 2, "Bold text must embed a different font program from regular text.");
    }

    private static HtmlPdfRenderRequestResult RenderWithWebFont(byte[] font, string text) {
        string html = "<style>@font-face{font-family:Scoped;src:url(data:font/ttf;base64," + Convert.ToBase64String(font)
            + ")}</style><p style='font-family:Scoped;font-size:40px'>" + text + "</p>";
        return HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions()));
    }
}
