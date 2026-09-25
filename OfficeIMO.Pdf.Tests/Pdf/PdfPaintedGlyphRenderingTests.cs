using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPaintedGlyphRenderingTests {
    [Fact]
    public void SubstituteEncodingPaintsItsGlyphWithoutReplacingToUnicodeSceneText() {
        const string content = "BT /F1 20 Tf 20 80 Td (A) Tj ET";
        const string cmap = "begincmap\n1 begincodespacerange\n<00> <FF>\nendcodespacerange\n1 beginbfchar\n<41> <0051>\nendbfchar\nendcmap";
        byte[] pdf = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>", "endobj",
            "4 0 obj", $"<< /Length {content.Length} >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding /Differences [65 /B] >> /ToUnicode 6 0 R >>", "endobj",
            "6 0 obj", $"<< /Length {cmap.Length} >>", "stream", cmap, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }) + "\n");

        OfficeDrawing drawing = PdfReadDocument.Open(pdf).Pages[0].ToDrawing();
        OfficeDrawingText visual = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());

        Assert.Equal("Q", visual.Text);
        Assert.Equal("B", visual.RasterText);
    }

    [Theory]
    [InlineData("Helvetica", "3 Tc", "B", "0051", "QQ")]
    [InlineData("Helvetica", "3 Tc", "B", "00510051", "QQQQ")]
    [InlineData("Symbol", "", "•", "0051", "QQ")]
    public void SubstitutedGlyphKeepsLogicalTextAcrossSpacingAndSymbolDifferences(
        string fontName, string spacing, string painted, string unicodeHex, string logicalText) {
        string content = $"BT /F1 20 Tf {spacing} 20 80 Td (AA) Tj ET";
        string cmap = $"begincmap\n1 begincodespacerange\n<00> <FF>\nendcodespacerange\n1 beginbfchar\n<41> <{unicodeHex}>\nendbfchar\nendcmap";
        string glyphName = fontName == "Symbol" ? "bullet" : "B";
        byte[] pdf = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>", "endobj",
            "4 0 obj", $"<< /Length {content.Length} >>", "stream", content, "endstream", "endobj",
            "5 0 obj", $"<< /Type /Font /Subtype /Type1 /BaseFont /{fontName} /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding /Differences [65 /{glyphName}] >> /ToUnicode 6 0 R >>", "endobj",
            "6 0 obj", $"<< /Length {cmap.Length} >>", "stream", cmap, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }) + "\n");

        OfficeDrawing drawing = PdfReadDocument.Open(pdf).Pages[0].ToDrawing();
        string logical = string.Concat(drawing.Elements.OfType<OfficeDrawingText>().Select(element => element.Text));
        string visual = string.Concat(drawing.Elements.OfType<OfficeDrawingText>().Select(element => element.RasterText));

        Assert.Equal(logicalText, logical);
        Assert.Equal(painted + painted, visual);
        if (spacing.Length > 0) Assert.Equal(2, drawing.Elements.OfType<OfficeDrawingText>().Count());
    }

    [Fact]
    public void PaintedGlyphMapOnlyNeedsRebuildWhenItsAliasesChangeForThatDrawing() {
        byte[] source = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B');
        var program = new PdfDrawingFontProgram(source, new SortedDictionary<int, int>(), _ => 2, _ => false);
        var aliases = new PdfReadPage.PaintedGlyphMap(program);
        var root = new OfficeDrawing(100, 100);
        var nested = new OfficeDrawing(100, 100);
        root.Fonts.Add("Test", source);
        nested.Fonts.Add("Test", source);

        aliases.Alias(2);
        Assert.True(aliases.NeedsApply(root));
        aliases.MarkApplied(root, ("Test", OfficeFontStyle.Regular));
        Assert.False(aliases.NeedsApply(root));
        root.Fonts.AddRange(nested.Fonts);
        Assert.True(aliases.NeedsApply(root));
        aliases.MarkApplied(root, ("Test", OfficeFontStyle.Regular));
        Assert.True(aliases.NeedsApply(nested));
        aliases.MarkApplied(nested, ("Test", OfficeFontStyle.Regular));
        aliases.Alias(3);
        Assert.True(aliases.NeedsApply(root));
        Assert.True(aliases.NeedsApply(nested));
    }

    [Fact]
    public void PaintedAliasUsesSupplementaryPrivateUseWhenBmpRangeIsClaimed() {
        byte[] source = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B');
        var claimed = new SortedDictionary<int, int>();
        for (int scalar = 0xE000; scalar <= 0xF8FF; scalar++) claimed.Add(scalar, 1);
        var program = new PdfDrawingFontProgram(source, claimed, _ => 2, _ => false);
        var aliases = new PdfReadPage.PaintedGlyphMap(program);

        int alias = aliases.Alias(2);
        Assert.Equal(0xF0000, alias);
        Assert.Equal(alias, aliases.Alias(2));
        Assert.Equal(2, aliases.Additions[alias]);
        byte[] rebuilt = Assert.IsType<byte[]>(PdfTrueTypeUnicodeCmap.TryAddMappings(program, aliases.Additions));
        Assert.True(OfficeTrueTypeFont.TryLoad(rebuilt)?.HasGlyphs(char.ConvertFromUtf32(alias)));

        PdfTextSpan visual = CreateGlyphRun("ffi", new[] { 3 }).WithVisualGlyph(alias);
        Assert.Equal(char.ConvertFromUtf32(alias), visual.Text);
        Assert.Equal("ffi", visual.LogicalDrawingText);
        Assert.Equal(new[] { 2 }, visual.GlyphCharacterLengths);
        Assert.Equal(2, visual.CharacterAdvances?.Count);
    }

    [Fact]
    public void VisualTextProjectionRetainsMatchingLogicalBreakProvenance() {
        var span = new PdfTextSpan("A B", "F1", 12, 10, 10, 24, null, true, 0, "Subset", null,
            embeddedLineBreakCounts: new[] { 0, 1, 0 });

        PdfTextSpan visual = span.WithVisualText("X Y");

        Assert.Equal("A B", visual.LogicalDrawingText);
        Assert.Equal(new[] { 0, 1, 0 }, visual.EmbeddedLineBreakCounts);
        Assert.Null(span.WithVisualText("glyph").EmbeddedLineBreakCounts);
    }

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
    public void InvisibleComplexRunDoesNotSpendPaintedGlyphBudget() {
        var span = new PdfTextSpan("ffiX", "F1", 12, 10, 10, 24, null, false, 0, "Subset", null,
            drawingFontFamily: "Subset", characterAdvances: [6D, 6D, 6D, 6D],
            glyphCharacterLengths: [3, 1], glyphBytes: [[65], [66]], glyphPaintedAdvances: [12D, 12D]);
        var spans = new List<PdfTextSpan> { span };

        PdfPaintedGlyphRuns.SplitComplexRuns(spans, _ => throw new InvalidOperationException("Invisible text was charged"));

        Assert.Same(span, Assert.Single(spans));
    }

    [Fact]
    public void FaintPaintedLigatureStillSplitsAtItsGlyphBoundary() {
        var span = new PdfTextSpan("ffiX", "F1", 12, 10, 10, 24,
            OfficeColor.FromRgba(0, 0, 0, 1), true, 0, "Subset", null,
            drawingFontFamily: "Subset", characterAdvances: [6D, 6D, 6D, 6D],
            glyphCharacterLengths: [3, 1], glyphBytes: [[65], [66]], glyphPaintedAdvances: [12D, 12D]);
        var spans = new List<PdfTextSpan> { span };
        int charged = 0;

        PdfPaintedGlyphRuns.SplitComplexRuns(spans, count => charged += count);

        Assert.Equal(2, charged);
        Assert.Equal(new[] { "ffi", "X" }, spans.Select(glyph => glyph.Text));
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

    [Fact]
    public void InkedNotdefGlyphCanTriggerPerGlyphVisualProjection() {
        PdfTextSpan span = CreateGlyphRun("AB", new[] { 1, 1 });
        var inked = new PdfDrawingFontProgram(Array.Empty<byte>(), new SortedDictionary<int, int>(), _ => 0, _ => false);
        var empty = new PdfDrawingFontProgram(Array.Empty<byte>(), new SortedDictionary<int, int>(), _ => 0, _ => true);

        Assert.Equal(2, Assert.IsType<List<PdfTextSpan>>(
            PdfPaintedGlyphRuns.SplitAlternateGlyphRun(span, inked, _ => { })).Count);
        Assert.Null(PdfPaintedGlyphRuns.SplitAlternateGlyphRun(span, empty, _ => { }));
    }

    [Fact]
    public void InkedNotdefAliasIsCoveredAndDrawnButAbsentScalarIsNot() {
        byte[] source = ManagedTextShapingTestAssets.CreateFontWithInkedNotdef();
        var program = new PdfDrawingFontProgram(source, new SortedDictionary<int, int> { ['A'] = 1 },
            _ => 0, glyph => glyph != 0);
        var aliases = new PdfReadPage.PaintedGlyphMap(program);
        int alias = aliases.Alias(0);
        byte[] rebuilt = Assert.IsType<byte[]>(PdfTrueTypeUnicodeCmap.TryAddMappings(program, aliases.Additions));
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(rebuilt));
        string painted = char.ConvertFromUtf32(alias);

        Assert.True(font.HasGlyphs(painted));
        Assert.NotEmpty(font.GetTextContours(painted, 0, 0, 12));
        Assert.False(font.HasGlyphs("\uE001"));

        // A regular font with the same cmap entry must treat glyph zero as missing.
        byte[] ordinary = (byte[])rebuilt.Clone();
        int tableCount = ordinary[4] << 8 | ordinary[5];
        for (int table = 0; table < tableCount; table++) {
            int record = 12 + table * 16;
            if (ordinary[record] == (byte)'p' && ordinary[record + 1] == (byte)'G' &&
                ordinary[record + 2] == (byte)'0' && ordinary[record + 3] == (byte)'0') {
                ordinary[record] = (byte)'x';
                break;
            }
        }
        Assert.False(Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(ordinary)).HasGlyphs(painted));
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
    public void SharedEmbeddedProgramWithDifferentPdfCodeMapsKeepsDistinctDrawingFaces() {
        byte[] bytes = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B');
        var source = new PdfFontResource("F1", "SharedBase", "WinAnsiEncoding", false,
            embeddedTrueTypeFont: bytes, fontSubtype: "TrueType", embeddedProgramSubtype: "TrueType");
        var unicode = new SortedDictionary<int, int> { ['A'] = 1, ['B'] = 2 };
        var first = new PdfDrawingFontProgram(bytes, unicode, code => code == 65 ? 1 : 0, _ => false);
        var second = new PdfDrawingFontProgram(bytes, unicode, code => code == 65 ? 2 : 0, _ => false);

        Assert.NotEqual(source.WithDrawingProgram(first).DrawingFontFamily,
            source.WithDrawingProgram(second).DrawingFontFamily);
        Assert.Equal(source.WithDrawingProgram(first).DrawingFontFamily,
            source.WithDrawingProgram(first).DrawingFontFamily);
    }

    [Fact]
    public void SharedCidProgramWithDifferentHighCidMappingsKeepsDistinctDrawingFaces() {
        byte[] bytes = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B');
        var source = new PdfFontResource("F1", "SharedCid", "Identity-H", true,
            embeddedTrueTypeFont: bytes, fontSubtype: "CIDFontType2", embeddedProgramSubtype: "TrueType");
        var unicode = new SortedDictionary<int, int> { ['A'] = 1 };
        byte[] firstMap = new byte[514];
        byte[] secondMap = new byte[514];
        firstMap[513] = 1;
        secondMap[513] = 2;
        var first = new PdfDrawingFontProgram(bytes, unicode, code => code == 256 ? 1 : 0,
            _ => false, firstMap);
        var second = new PdfDrawingFontProgram(bytes, unicode, code => code == 256 ? 2 : 0,
            _ => false, secondMap);

        Assert.NotEqual(source.WithDrawingProgram(first).DrawingFontFamily,
            source.WithDrawingProgram(second).DrawingFontFamily);
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
