using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlPdfColorEmojiTests {
    private const string Text = "\U0001F600\U0001F600";

    // HTML paints color fonts as vector outlines: COLR layers as filled paths, PNG glyphs as drawing images.
    [Theory]
    [InlineData("colr")]
    [InlineData("cbdt")]
    [InlineData("sbix")]
    public void HtmlColorGlyphIsPaintedInColorWithItsText(string kind) {
        byte[] font = kind switch {
            "colr" => ManagedTextShapingTestAssets.CreateColorFont(0x1F600),
            "cbdt" => ManagedTextShapingTestAssets.CreateBitmapColorFont(0x1F600, sbix: false),
            _ => ManagedTextShapingTestAssets.CreateBitmapColorFont(0x1F600, sbix: true)
        };

        byte[] pdf = ConvertHtml(font, Text);

        Assert.Contains(Text, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(CountRed(pdf) > 0, "The color glyph artwork is painted.");
    }

    // A color font's sequence becomes its color glyph only through the font's ccmp ligature, while its
    // components are plain. Every fixture glyph is a box narrower than its advance and the joiner is
    // inked, so unshaped components paint three black boxes instead of one red one.
    [Fact]
    public void HtmlColorEmojiSequenceWithoutFontFeaturesPaintsItsColorLigature() {
        const string sequence = "\U0001F3CB\u200D\u2642";

        byte[] pdf = ConvertHtml(ManagedTextShapingTestAssets.CreateEmojiSequenceFont(redFirstLigature: true), sequence);

        Assert.Contains(sequence, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(CountRed(pdf) > 0, "The ligature's red layer is painted.");
        Assert.Equal(1, CountInkSegmentsOnTheDensestRow(pdf));
    }

    // Glyph artwork keeps source paint order: a letter after a PNG glyph and overlapping it stays on top.
    [Fact]
    public void LetterAfterAPngGlyphIsPaintedOverIt() {
        byte[] font = ManagedTextShapingTestAssets.CreateBitmapColorFont(0x1F600, sbix: false, overlappingLetter: 'A');

        int alone = CountDark(ConvertHtml(font, "A"));
        int overlapping = CountDark(ConvertHtml(font, "\U0001F600A"));

        Assert.True(alone > 0, "The letter is painted.");
        Assert.InRange(overlapping, alone * 0.9D, alone * 1.1D);
    }

    [Fact]
    public void MarkdownPngGlyphIsPaintedWithItsText() {
        // Markdown text uses the default standard font; the fixture stands in for it, without fallbacks.
        var options = new PdfCore.PdfCreateOptions { IncludeStandardFontToUnicodeMaps = true, DefaultFont = PdfCore.PdfStandardFont.Helvetica }
            .EmbedStandardFont(PdfCore.PdfStandardFont.Helvetica, ManagedTextShapingTestAssets.CreateBitmapColorFont(0x1F600, sbix: false));

        byte[] pdf = MarkdownDocument.Parse(Text).ToPdf(new MarkdownPdfExportOptions {
            PdfOptions = options, TextFallbacks = PdfCore.PdfTextFallbackFeatures.None, ShowPageNumbers = false
        }).ToBytes();

        Assert.Contains(Text, PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(CountRed(pdf) > 0, "The PNG glyph is painted.");
    }

    private static byte[] ConvertHtml(byte[] font, string text) {
        string html = "<style>@font-face{font-family:Emoji;src:url(data:font/ttf;base64," + Convert.ToBase64String(font)
            + ")}p{font:60px Emoji}</style><p>" + text + "</p>";
        return HtmlConversionDocument.Parse(html).ConvertToPdfWithEvidence(new HtmlPdfExportOptions()).ToBytes();
    }

    private static OfficeRasterImage Render(byte[] pdf) =>
        OfficeRasterImageDecoder.Decode(PdfCore.PdfPageImageRenderer.RenderPageAsPng(pdf, 1, 1D));

    private static int CountRed(byte[] pdf) {
        OfficeRasterImage image = Render(pdf);
        int red = 0;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor color = image.GetPixel(x, y);
                if (color.R > 200 && color.G < 60 && color.B < 60) red++;
            }
        }
        return red;
    }

    private static int CountDark(byte[] pdf) {
        OfficeRasterImage image = Render(pdf);
        int dark = 0;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor color = image.GetPixel(x, y);
                if (color.R < 80 && color.G < 80 && color.B < 80) dark++;
            }
        }
        return dark;
    }

    private static int CountInkSegmentsOnTheDensestRow(byte[] pdf) {
        OfficeRasterImage image = Render(pdf);
        int bestInk = -1, bestSegments = 0;
        for (int y = 0; y < image.Height; y++) {
            int ink = 0, segments = 0;
            bool inside = false;
            for (int x = 0; x < image.Width; x++) {
                OfficeColor color = image.GetPixel(x, y);
                // Ink is black or the red layer: anything with little green and blue.
                bool dark = color.G < 128 && color.B < 128;
                if (dark) ink++;
                if (dark && !inside) segments++;
                inside = dark;
            }
            if (ink > bestInk) { bestInk = ink; bestSegments = segments; }
        }
        return bestSegments;
    }
}
