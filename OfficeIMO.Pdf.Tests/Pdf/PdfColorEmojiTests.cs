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

    // A rotated text matrix carries the PNG with it: the 0.8 em square keeps its area, and turned by
    // 45 degrees its bounding box grows by the square root of two. A text watermark is rotated text.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PngGlyphInRotatedWatermarkTextIsPaintedAsARotatedSquare(bool sbix) {
        byte[] pdf = Save(sbix ? "sbix" : "cbdt", (content, _) => content.AddParagraph(paragraph =>
            paragraph.AddRuns(new[] { PdfTextRun.Normal("Body", standardFont: PdfStandardFont.Courier) })),
            options => options.TextWatermark = new PdfTextWatermark("\U0001F600") {
                Font = PdfStandardFont.Helvetica, Bold = false, Opacity = 1D, RotationAngle = 45D, FontSize = OfficeLength.Points(FontSize)
            });

        Pixels pixels = CountPixels(pdf);
        double side = 0.8D * FontSize;
        Assert.InRange(pixels.Red, side * side * 0.9D, side * side * 1.1D);
        Assert.InRange(pixels.RedWidth, side * Math.Sqrt(2D) - 3D, side * Math.Sqrt(2D) + 3D);
        Assert.InRange(pixels.RedHeight, side * Math.Sqrt(2D) - 3D, side * Math.Sqrt(2D) + 3D);
    }

    // Synthetic oblique shears the text matrix by one third, and the PNG with it: the image's unit
    // square maps to a 0.8 em square whose top edge moves right by a third of its height.
    [Fact]
    public void PngGlyphInSyntheticObliqueTextIsShearedWithTheText() {
        byte[] pdf = Save("cbdt", (content, _) => content.AddParagraph(paragraph =>
            paragraph.AddRuns(new[] { PdfTextRun.Italicized("\U0001F600", fontFamily: Family) })));

        System.Text.RegularExpressions.Match draw = System.Text.RegularExpressions.Regex.Match(ContentText(pdf),
            @"(-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) -?[\d.]+ -?[\d.]+ cm\s+/Im\d+ Do");
        Assert.True(draw.Success, "The PNG glyph is drawn as an image.");
        double[] matrix = Enumerable.Range(1, 4).Select(group => double.Parse(draw.Groups[group].Value, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
        double side = 0.8D * FontSize;
        Assert.Equal(side, matrix[0], 2);
        Assert.Equal(0D, matrix[1], 2);
        // The text matrix is written with the shear rounded to 0.333.
        Assert.InRange(matrix[2], side / 3D - 0.05D, side / 3D + 0.05D);
        Assert.Equal(side, matrix[3], 2);
    }

    // A translucent palette color is painted through an ExtGState fill alpha, one state per alpha value.
    [Fact]
    public void TranslucentPaletteColorIsPaintedWithItsAlphaThroughOneSharedState() {
        byte[] pdf = Save("half-red", (content, _) => content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal(Text) })));

        Pixels pixels = CountPixels(pdf);
        Assert.Equal(0, pixels.Red);
        Assert.True(pixels.HalfRed > 0, "The half-transparent red layer blends with the white page.");
        string[] states = System.Text.RegularExpressions.Regex.Matches(ContentText(pdf), @"/(GS\d+) gs")
            .Cast<System.Text.RegularExpressions.Match>().Select(match => match.Groups[1].Value).ToArray();
        Assert.Equal(2, states.Length);
        Assert.Single(states.Distinct());
    }

    // A text-color layer with paint alpha keeps that alpha over the opaque text fill.
    [Fact]
    public void TextColorLayerKeepsItsPaintAlpha() {
        byte[] pdf = Save("half-foreground", (content, _) => content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal("\U0001F600") })));

        Pixels pixels = CountPixels(pdf);
        Assert.Equal(0, pixels.Dark);
        Assert.True(pixels.Gray > 0, "The half-transparent text-color layer blends with the white page.");
    }

    // COLR artwork whose layers are all transparent paints nothing; the inked base outline stays hidden.
    [Fact]
    public void TransparentColrArtworkPaintsNothing() {
        byte[] pdf = Save("transparent", (content, _) => content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal(Text) })));

        Assert.Equal(Text, PdfReadDocument.Open(pdf).ExtractText().Trim());
        Pixels pixels = CountPixels(pdf);
        Assert.Equal(0, pixels.Dark + pixels.Red + pixels.Blue);
    }

    // A bidi run is written in segments; its color glyphs are still painted at their visual positions.
    [Fact]
    public void ColorGlyphInAMixedDirectionRunIsPainted() {
        const string text = "\u05D0\U0001F600";
        byte[] pdf = Save("cbdt-hebrew", (content, _) => content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal(text) })));

        Assert.Equal(text, PdfReadDocument.Open(pdf).ExtractText().Trim());
        // Two 0.8 em squares one 0.5 em advance apart: their union, without the overlap.
        double side = 0.8D * FontSize, overlap = side - 0.5D * FontSize;
        double union = 2D * side * side - overlap * side;
        Assert.InRange(CountPixels(pdf).Red, union * 0.93D, union * 1.07D);
    }

    // Ordinary glyphs in a synthetic-bold run with a color glyph keep their synthetic stroke.
    [Fact]
    public void SyntheticBoldLettersBesideAColorGlyphKeepTheirWeight() {
        const string sequence = "\U0001F3CB\u200D\u2642";
        int Dark(Func<string, PdfTextRun> run, string text) =>
            CountPixels(Save("red-sequence", (content, _) => content.AddParagraph(paragraph => paragraph.AddRuns(new[] { run(text) })))).Dark;

        int boldAlone = Dark(text => PdfTextRun.Bolded(text, fontFamily: Family), "A");
        int regularAlone = Dark(text => PdfTextRun.Normal(text, fontFamily: Family), "A");
        int boldBesideColor = Dark(text => PdfTextRun.Bolded(text, fontFamily: Family), "A" + sequence);

        Assert.True(boldAlone > regularAlone * 1.05D, "Synthetic bold widens the letter.");
        Assert.InRange(boldBesideColor, boldAlone * 0.97D, boldAlone * 1.03D);
    }

    // The synthetic-weight scope's Q restores the rendering mode set inside it; a later color run in the
    // same builder must return to fill, not to the scoped fill-and-stroke mode.
    [Fact]
    public void ColorRunAfterASyntheticBoldTextObjectRestoresFillRendering() {
        PdfTrueTypeFontProgram font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateColorFont(0x1F600), "Color");
        PdfGlyphRun run = font.ShapeText("\U0001F600", PdfTextShapingOptions.ForRendering(font.FontName, PdfTextShapingMode.OpenTypeLigatures));
        PdfTextShowCommand command = font.ToTextShowCommand("\U0001F600", run);
        var content = new StringBuilder();
        var builder = new ContentStreamBuilder(content);

        builder.BeginText().TextMatrix(0D, 0D).Font("F1", 12D, syntheticBold: true).ShowText(command, 12D).EndText();
        int second = content.Length;
        builder.BeginText().TextMatrix(0D, 20D).Font("F1", 12D).ShowText(command, 12D).EndText();

        string colorRun = content.ToString(second, content.Length - second);
        Assert.Contains("3 Tr", colorRun, StringComparison.Ordinal);
        Assert.DoesNotContain("2 Tr", colorRun, StringComparison.Ordinal);
    }

    private static byte[] Save(string kind, Action<PdfContentBuilder, PdfPageBuilder> build, Action<PdfCreateOptions>? configure = null) {
        PdfCreateOptions options = CreateOptions(kind);
        configure?.Invoke(options);
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
            // Half-transparent red under a fully transparent blue layer.
            "half-red" => ManagedTextShapingTestAssets.CreateColorFontWithPaletteSize(0x1F600, 2, redAlpha: 128, blueAlpha: 0),
            // Both layers fully transparent over an inked base outline.
            "transparent" => ManagedTextShapingTestAssets.CreateColorFontWithPaletteSize(0x1F600, 2, redAlpha: 0, blueAlpha: 0),
            "half-foreground" => ManagedTextShapingTestAssets.CreateHalfAlphaForegroundColrFont(0x1F600),
            "cbdt-hebrew" => ManagedTextShapingTestAssets.CreateBitmapColorFont(0x1F600, sbix: false, sameGlyphScalars: new[] { 0x05D0 }),
            "red-sequence" => ManagedTextShapingTestAssets.CreateEmojiSequenceFont(redFirstLigature: true),
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

    private readonly record struct Pixels(int Red, int Blue, int HalfRed, int RedWidth, int RedHeight, int Dark, int Gray);

    private static Pixels CountPixels(byte[] pdf) {
        OfficeRasterImage image = OfficeRasterImageDecoder.Decode(PdfPageImageRenderer.RenderPageAsPng(pdf, 1, 1D));
        int red = 0, blue = 0, halfRed = 0, dark = 0, gray = 0, left = int.MaxValue, right = -1, top = int.MaxValue, bottom = -1;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor color = image.GetPixel(x, y);
                if (color.R > 200 && color.G < 60 && color.B < 60) {
                    red++;
                    left = Math.Min(left, x); right = Math.Max(right, x);
                    top = Math.Min(top, y); bottom = Math.Max(bottom, y);
                } else if (color.B > 200 && color.R < 60 && color.G < 60) {
                    blue++;
                } else if (color.R > 200 && color.G is > 90 and < 170 && color.B is > 90 and < 170) {
                    halfRed++;
                } else if (color.R < 80 && color.G < 80 && color.B < 80) {
                    dark++;
                } else if (color.R is > 90 and < 190 && Math.Abs(color.R - color.G) < 20 && Math.Abs(color.G - color.B) < 30) {
                    gray++;
                }
            }
        }
        return new Pixels(red, blue, halfRed, red == 0 ? 0 : right - left + 1, red == 0 ? 0 : bottom - top + 1, dark, gray);
    }
}
