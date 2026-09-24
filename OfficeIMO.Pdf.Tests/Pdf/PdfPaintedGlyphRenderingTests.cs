using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPaintedGlyphRenderingTests {
    [Fact]
    public void ComplexRunChargesExpansionBeforeReplacingItsSourceSpan() {
        PdfTextSpan span = CreateGlyphRun("ffiX", new[] { 3, 1 });
        var spans = new List<PdfTextSpan> { span };

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfPaintedGlyphRuns.SplitComplexRuns(spans, count => {
                Assert.Equal(2, count);
                throw new InvalidOperationException("budget exceeded");
            }));

        Assert.Equal("budget exceeded", exception.Message);
        Assert.Same(span, Assert.Single(spans));
    }

    [Fact]
    public void AlternateRunChargesExpansionBeforeCreatingGlyphSpans() {
        PdfTextSpan span = CreateGlyphRun("AB", new[] { 1, 1 });
        var program = new PdfDrawingFontProgram(Array.Empty<byte>(), new SortedDictionary<int, int> {
            ['A'] = 2, ['B'] = 3
        }, _ => 4, _ => false);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfPaintedGlyphRuns.SplitAlternateGlyphRun(span, program, count => {
                Assert.Equal(2, count);
                throw new InvalidOperationException("budget exceeded");
            }));

        Assert.Equal("budget exceeded", exception.Message);
    }

    private static PdfTextSpan CreateGlyphRun(string text, int[] glyphLengths) => new(
        text, "F1", 12, 10, 10, 24, null, true, 0, "Subset", null,
        drawingFontFamily: "Subset", characterAdvances: Enumerable.Repeat(6D, text.Length).ToArray(),
        glyphCharacterLengths: glyphLengths,
        glyphBytes: Enumerable.Range(0, glyphLengths.Length).Select(index => new[] { (byte)('A' + index) }).ToArray(),
        glyphPaintedAdvances: Enumerable.Repeat(12D, glyphLengths.Length).ToArray());

    [Fact]
    public void SupplementaryScalarWithDifferentPaintedGlyphSplitsAtScalarBoundary() {
        string value = char.ConvertFromUtf32(0x1F600) + "A";
        PdfTextSpan span = CreateGlyphRun(value, new[] { 2, 1 });
        var program = new PdfDrawingFontProgram(Array.Empty<byte>(), new SortedDictionary<int, int> {
            [0x1F600] = 2, ['A'] = 3
        }, _ => 4, _ => false);

        List<PdfTextSpan> glyphs = Assert.IsType<List<PdfTextSpan>>(
            PdfPaintedGlyphRuns.SplitAlternateGlyphRun(span, program, _ => { }));
        Assert.Equal(2, glyphs.Count);
        Assert.Equal(char.ConvertFromUtf32(0x1F600), glyphs[0].Text);
        Assert.Equal("A", glyphs[1].Text);
    }

    [Fact]
    public void LigatureOnlySubsetRegistersItsPaintedGlyph() {
        string root = VisualBaselineTestSupport.GetTestsProjectRoot();
        string path = Path.Combine(root, "Pdf", "Fixtures", "ShapedText", "ligature-only.pdf");
        OfficeDrawing drawing = PdfReadDocument.Open(File.ReadAllBytes(path)).Pages[0].ToDrawing();

        OfficeDrawingText visual = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("ffi", visual.Text);
        Assert.Single(visual.RasterText);
        Assert.InRange(visual.RasterText[0], '\uE000', '\uF8FF');
        Assert.Equal(visual.Text, visual.Clone().Text);
        Assert.Equal(visual.RasterText, visual.Clone().RasterText);
        Assert.Contains(visual.RasterText, OfficeDrawingSvgExporter.ToSvg(drawing));
        Assert.NotEmpty(drawing.Fonts.Faces);
    }

    [Fact]
    public void CallerSuppliedFaceRemainsSelectedWhenPaintedGlyphNeedsAnAlias() {
        string root = VisualBaselineTestSupport.GetTestsProjectRoot();
        string path = Path.Combine(root, "Pdf", "Fixtures", "ShapedText", "ligature-only.pdf");
        PdfReadPage page = PdfReadDocument.Open(File.ReadAllBytes(path)).Pages[0];
        OfficeFontFace embedded = Assert.Single(page.ToDrawing().Fonts.Faces);
        byte[] replacement = ManagedTextShapingTestAssets.CreateFont(' ', 'A', 'B');
        var fonts = new OfficeFontFaceCollection().Add(embedded.FamilyName, replacement, embedded.Style);

        OfficeDrawing configured = page.ToDrawing(fonts);

        OfficeFontFace face = Assert.Single(configured.Fonts.Faces);
        Assert.Equal(replacement, face.Data);
    }

    [Fact]
    public void FullEmbeddedTrueTypeFontReceivesAnIsolatedDrawingFamily() {
        byte[] program = ManagedTextShapingTestAssets.CreateFont(' ', 'A');
        var resource = new PdfFontResource("F1", "FullFont", "WinAnsiEncoding", false,
            embeddedTrueTypeFont: program, fontSubtype: "TrueType", embeddedProgramSubtype: "TrueType");
        var drawingProgram = new PdfDrawingFontProgram(program, new SortedDictionary<int, int>(),
            _ => 0, _ => false);

        Assert.StartsWith("FullFont-", resource.DrawingFontFamily, StringComparison.Ordinal);
        Assert.StartsWith("FullFont-", resource.WithDrawingProgram(drawingProgram).DrawingFontFamily,
            StringComparison.Ordinal);
    }

    [Fact]
    public void FullEmbeddedFontsWithTheSamePdfNameKeepDistinctDrawingFamilies() {
        var pageFont = new PdfFontResource("F1", "SharedBase", "WinAnsiEncoding", false,
            embeddedTrueTypeFont: ManagedTextShapingTestAssets.CreateFont(' ', 'A'));
        var annotationFont = new PdfFontResource("F2", "SharedBase", "WinAnsiEncoding", false,
            embeddedTrueTypeFont: ManagedTextShapingTestAssets.CreateFont(' ', 'B'));

        Assert.NotEqual(pageFont.DrawingFontFamily, annotationFont.DrawingFontFamily);
    }

    [Fact]
    public void SimpleSymbolicSubsetWithOnlyClusterMappingsRegistersItsPaintedGlyphs() {
        string root = VisualBaselineTestSupport.GetTestsProjectRoot();
        string path = Path.Combine(root, "Pdf", "Fixtures", "Fonts", "symbolic-truetype-cluster-only.pdf");
        OfficeDrawing drawing = PdfReadDocument.Open(File.ReadAllBytes(path)).Pages[0].ToDrawing();

        Assert.NotEmpty(drawing.Fonts.Faces);
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), visual =>
            visual.RasterText.Any(character => character >= '\uE000' && character <= '\uF8FF') &&
            !visual.Text.Any(character => character >= '\uE000' && character <= '\uF8FF'));
    }

    [Fact]
    public void SimpleSymbolicSubsetUsesPaintedCodeWhenItsUnicodeCmapDisagrees() {
        string root = VisualBaselineTestSupport.GetTestsProjectRoot();
        string path = Path.Combine(root, "Pdf", "Fixtures", "Fonts", "symbolic-truetype-conflicting-unicode.pdf");
        OfficeDrawing drawing = PdfReadDocument.Open(File.ReadAllBytes(path)).Pages[0].ToDrawing();

        OfficeFontFace face = Assert.Single(drawing.Fonts.Faces);
        Assert.True(face.Program.TryGetGlyphMetrics('S', out int paintedGlyph, out _));
        Assert.True(face.Program.TryGetGlyphMetrics(' ', out int emptyGlyph, out _));
        Assert.NotEqual(emptyGlyph, paintedGlyph);
    }

    [Theory]
    [InlineData("empty-glyph-text-only.pdf")]
    [InlineData("inked-space-only.pdf")]
    public void CidSubsetWithOnlyUnusableUnicodeTextStillRegistersPaintedGlyph(string fileName) {
        string root = VisualBaselineTestSupport.GetTestsProjectRoot();
        string path = Path.Combine(root, "Pdf", "Fixtures", "ShapedText", fileName);
        OfficeDrawing drawing = PdfReadDocument.Open(File.ReadAllBytes(path)).Pages[0].ToDrawing();

        Assert.NotEmpty(drawing.Fonts.Faces);
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), visual =>
            visual.RasterText.Any(character => character >= '\uE000' && character <= '\uF8FF') &&
            !visual.Text.Any(character => character >= '\uE000' && character <= '\uF8FF'));
    }

    // A PDF paints shaped, positioned glyphs. Rendering must reproduce those exact glyphs: contextual
    // Arabic, Persian and Urdu forms, font-specific alternates, words painted across two fonts,
    // clipped glyphs at a line edge, rotated right-to-left lines, Devanagari clusters and glyphs whose
    // ToUnicode text is U+0000, inked glyphs coded as spaces (colour-font layers), ligature glyphs inside
    // multi-glyph runs (including Ghostscript's mismatched ligature text), simple fonts keyed by
    // character code, Word runs whose trailing spaces are painted inside TJ, and repeated spaces.
    // The oracle is Poppler, an independent renderer; references come from create_references.py.
    [Theory]
    [InlineData("chrome-arabic", "Pdf/Fixtures/ShapedText/chrome-arabic.pdf")]
    [InlineData("chrome-arabic-extended", "Pdf/Fixtures/ShapedText/chrome-arabic-extended.pdf")]
    [InlineData("cairo-arabic-extended", "Pdf/Fixtures/ShapedText/cairo-arabic-extended.pdf")]
    [InlineData("chrome-devanagari", "Pdf/Fixtures/ShapedText/chrome-devanagari.pdf")]
    [InlineData("space-coded-glyph", "Pdf/Fixtures/ShapedText/space-coded-glyph.pdf")]
    [InlineData("ligature-run", "Pdf/Fixtures/ShapedText/ligature-run.pdf")]
    [InlineData("ligature-only", "Pdf/Fixtures/ShapedText/ligature-only.pdf")]
    [InlineData("ghostscript-ligature-run", "Pdf/Fixtures/ShapedText/ghostscript-ligature-run.pdf")]
    [InlineData("cairo-rtl-0", "../OfficeIMO.TestAssets/MultilingualLayout/rtl-0-native.pdf")]
    [InlineData("cairo-rtl-90", "../OfficeIMO.TestAssets/MultilingualLayout/rtl-90-native.pdf")]
    [InlineData("cairo-latin-0", "../OfficeIMO.TestAssets/MultilingualLayout/latin-0-native.pdf")]
    [InlineData("word-mac-report", "Pdf/ReferenceBaselines/microsoft-word-16.109-native-word-report.pdf")]
    [InlineData("word-windows-summary", "Pdf/ReferenceBaselines/microsoft-word-windows-word-business-delivery-summary.pdf")]
    public void EmbeddedGlyphRunsMatchIndependentPopplerRaster(string reference, string relativePdfPath) {
        string root = VisualBaselineTestSupport.GetTestsProjectRoot();
        string pdfPath = Path.GetFullPath(Path.Combine(root, relativePdfPath));
        string referencePath = Path.Combine(root, "Pdf", "Fixtures", "ShapedText", "poppler-" + reference + ".png");

        byte[] rendered = PdfDocument.Load(pdfPath).Render.Pages("1", new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Png,
            Scale = 2D,
            MaxPages = 1
        })[0].Bytes!;
        OfficeRasterImage actual = VisualBaselineTestSupport.DecodePng(rendered, "OfficeIMO page raster is not a supported PNG file.");
        OfficeRasterImage expected = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(referencePath), "Poppler reference is not a supported PNG file.");
        Assert.InRange(Math.Abs(actual.Width - expected.Width), 0, 1);
        Assert.InRange(Math.Abs(actual.Height - expected.Height), 0, 1);

        int width = Math.Min(actual.Width, expected.Width);
        int height = Math.Min(actual.Height, expected.Height);
        bool[,] actualInk = Ink(actual, width, height);
        bool[,] expectedInk = Ink(expected, width, height);
        double precision = Coverage(actualInk, expectedInk);
        double recall = Coverage(expectedInk, actualInk);
        Assert.True(precision >= 0.99D && recall >= 0.99D,
            $"{reference}: ink precision {precision:0.000} and recall {recall:0.000} against Poppler must both be at least 0.990.");
    }

    private static bool[,] Ink(OfficeRasterImage image, int width, int height) {
        var ink = new bool[width, height];
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                double alpha = pixel.A / 255D;
                // Composite over white, then threshold luminance like the reference.
                double luminance = (0.299D * pixel.R + 0.587D * pixel.G + 0.114D * pixel.B) * alpha + 255D * (1D - alpha);
                ink[x, y] = luminance < 128D;
            }
        }
        return ink;
    }

    // Fraction of source ink pixels that have target ink within one pixel.
    private static double Coverage(bool[,] source, bool[,] target) {
        int width = source.GetLength(0);
        int height = source.GetLength(1);
        long total = 0;
        long covered = 0;
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                if (!source[x, y]) continue;
                total++;
                bool found = false;
                for (int dy = -1; dy <= 1 && !found; dy++) {
                    for (int dx = -1; dx <= 1 && !found; dx++) {
                        int nx = x + dx;
                        int ny = y + dy;
                        found = nx >= 0 && ny >= 0 && nx < width && ny < height && target[nx, ny];
                    }
                }
                if (found) covered++;
            }
        }
        return total == 0 ? 1D : covered / (double)total;
    }
}
