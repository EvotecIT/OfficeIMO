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
    public void HtmlPdf_UncoveredUnicodeFailsWithStructuredDiagnosticBeforeReturningDocument(
        HtmlRenderIntentProfile profile) {
        const string html = "<p>A−B</p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var request = HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });

        HtmlConversionException failure = Assert.Throws<HtmlConversionException>(() =>
            document.RenderToPdfDocumentResult(request));

        HtmlDiagnostic diagnostic = Assert.Single(failure.Diagnostics,
            item => item.Code == "unsupported-text-glyph");
        Assert.Equal(HtmlDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.Equal(OfficeConversionLossKind.Failure, diagnostic.LossKind);
        Assert.Contains("U+2212", diagnostic.Message, StringComparison.Ordinal);
        Assert.Contains("PdfCanvas", diagnostic.Detail, StringComparison.Ordinal);
        Assert.Equal(html, document.SourceHtml);
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

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);

        Assert.Contains("A−B", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(PdfCore.PdfDiagnostics.Analyze(pdf).EmbeddedFontCount >= 1);
    }
}
