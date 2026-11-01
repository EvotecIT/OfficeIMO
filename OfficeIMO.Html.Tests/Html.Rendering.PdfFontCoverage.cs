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
}
