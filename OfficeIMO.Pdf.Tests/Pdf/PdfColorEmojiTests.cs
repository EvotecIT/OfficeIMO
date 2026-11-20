using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

// PDF text embeds outlines only. A run with COLR or PNG color glyphs is written as invisible,
// extractable text, and each glyph's artwork is painted after the text object at the same pen.
public class PdfColorEmojiTests {
    // The fixtures map only the emoji; two of them check that the second is painted one advance later.
    private const string Text = "\U0001F600\U0001F600";
    private const string Family = "Color emoji";
    private const double FontSize = 50D;

    public static TheoryData<string> ColorFonts => new() { "colr", "large-palette", "cbdt", "sbix" };

    [Theory]
    [MemberData(nameof(ColorFonts))]
    public void ColorGlyphIsPaintedInColorOverExactlyExtractableText(string kind) {
        byte[] pdf = Save(kind, (content, _) => content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal(Text) })));

        Assert.Contains("3 Tr", ContentText(pdf), StringComparison.Ordinal);
        Assert.Equal(Text, PdfReadDocument.Open(pdf).ExtractText().Trim());
        Pixels pixels = CountPixels(pdf);
        Assert.True(pixels.Red > 0, "The red artwork is painted.");
        // COLR fixtures layer blue over red; the PNG fixtures are solid red.
        Assert.Equal(kind is "colr" or "large-palette", pixels.Blue > 0);
    }

    // The PNG is 80 pixels square at 100 ppem with a 1000-unit em: 0.8 em tall, sitting on the baseline.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PngGlyphIsPlacedAtItsStrikeSize(bool sbix) {
        byte[] pdf = Save(sbix ? "sbix" : "cbdt", (content, _) => content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal("\U0001F600") })));

        Pixels pixels = CountPixels(pdf);
        Assert.InRange(pixels.RedHeight, 0.8D * FontSize - 2D, 0.8D * FontSize + 2D);
        Assert.InRange(pixels.RedWidth, 0.8D * FontSize - 2D, 0.8D * FontSize + 2D);
    }

    public static TheoryData<string, string> PngLocations => new() {
        { "cbdt", "table" }, { "cbdt", "header" }, { "cbdt", "footer" },
        { "sbix", "table" }, { "sbix", "header" }, { "sbix", "footer" }
    };

    [Theory]
    [MemberData(nameof(PngLocations))]
    public void PngGlyphIsPaintedInTablesAndPageHeaders(string kind, string location) {
        byte[] pdf = Save(kind, (content, page) => {
            if (location == "table") {
                content.Item(item => item.AddTable(new[] { new[] { PdfTableCell.RichTextCell(new[] { PdfTextRun.Normal(Text, fontFamily: Family) }) } }, OfficeAlignment.Left));
                return;
            }
            // Courier body text has no red ink, so any red comes from the header or footer.
            content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal("Body", standardFont: PdfStandardFont.Courier) }));
            if (location == "header") page.Header(header => header.StandardFont(PdfStandardFont.Helvetica).FontSize(OfficeLength.Points(FontSize)).Text(Text));
            else page.Footer(footer => footer.StandardFont(PdfStandardFont.Helvetica).FontSize(OfficeLength.Points(FontSize)).Text(Text));
        });

        Assert.Contains(Text, PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(CountPixels(pdf).Red > 0, "The PNG glyph is painted.");
    }

    // A watermark is drawn into a Form, which has no page to place a PNG on: the preflight reports
    // it, while the same glyph in the page body is admitted.
    [Fact]
    public void PngGlyphInAWatermarkIsReportedInsteadOfWrittenBlank() {
        PdfCreateOptions options = CreateOptions("cbdt");
        options.TextWatermark = new PdfTextWatermark(Text) { Font = PdfStandardFont.Helvetica, Bold = false };
        PdfDocument document = PdfDocument.Create(doc => doc.Page(page => page.Content(content =>
            content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal(Text) })))), options);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = document.AnalyzeTextEncoding();
        // Each watermark variant that shows this text reports it.
        Assert.Equal(new[] { 0, 2 }, diagnostics.Select(diagnostic => diagnostic.Index).Distinct());
        Assert.All(diagnostics, diagnostic => Assert.Contains("PdfTextWatermark", diagnostic.Location, StringComparison.Ordinal));
    }

    private static byte[] Save(string kind, Action<PdfContentBuilder, PdfPageBuilder> build) {
        PdfCreateOptions options = CreateOptions(kind);
        PdfDocument document = PdfDocument.Create(doc => doc.Page(page => page.Content(content => build(content, page))), options);
        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = document.AnalyzeTextEncoding();
        Assert.True(diagnostics.Count == 0, string.Join(Environment.NewLine, diagnostics.Select(diagnostic => diagnostic.Message)));
        return document.ToBytes();
    }

    // The color font stands in for Helvetica, so body, table, header and footer text all use it.
    private static PdfCreateOptions CreateOptions(string kind) {
        byte[] font = kind switch {
            "colr" => ManagedTextShapingTestAssets.CreateColorFont(0x1F600),
            "large-palette" => ManagedTextShapingTestAssets.CreateColorFontWithPaletteSize(0x1F600, 5000),
            "cbdt" => ManagedTextShapingTestAssets.CreateBitmapColorFont(0x1F600, sbix: false),
            _ => ManagedTextShapingTestAssets.CreateBitmapColorFont(0x1F600, sbix: true)
        };
        return new PdfCreateOptions {
            IncludeStandardFontToUnicodeMaps = true,
            CompressContentStreams = false,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = OfficeLength.Points(FontSize)
        }.EmbedStandardFont(PdfStandardFont.Helvetica, font)
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(Family, font));
    }

    private static string ContentText(byte[] pdf) => Encoding.GetEncoding(28591).GetString(pdf);

    private readonly record struct Pixels(int Red, int Blue, int RedWidth, int RedHeight);

    private static Pixels CountPixels(byte[] pdf) {
        OfficeRasterImage image = OfficeRasterImageDecoder.Decode(PdfPageImageRenderer.RenderPageAsPng(pdf, 1, 1D));
        int red = 0, blue = 0, left = int.MaxValue, right = -1, top = int.MaxValue, bottom = -1;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor color = image.GetPixel(x, y);
                if (color.R > 200 && color.G < 60 && color.B < 60) {
                    red++;
                    left = Math.Min(left, x); right = Math.Max(right, x);
                    top = Math.Min(top, y); bottom = Math.Max(bottom, y);
                } else if (color.B > 200 && color.R < 60 && color.G < 60) {
                    blue++;
                }
            }
        }
        return new Pixels(red, blue, red == 0 ? 0 : right - left + 1, red == 0 ? 0 : bottom - top + 1);
    }
}
