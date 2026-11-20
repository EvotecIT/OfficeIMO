using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfEmojiSequenceTests {
    private const string Family = "Emoji sequences";
    private const string LiftingWeightsMale = "\U0001F3CB\uFE0F\u200D\u2642\uFE0F";

    public static TheoryData<string, int> Sequences => new() {
        { LiftingWeightsMale, ManagedTextShapingTestAssets.EmojiSequenceLigatures[0].Ligature },
        { "\U0001F3CB\u200D\u2642", ManagedTextShapingTestAssets.EmojiSequenceLigatures[0].Ligature },
        { "\U0001F44D\U0001F3FB", ManagedTextShapingTestAssets.EmojiSequenceLigatures[1].Ligature },
        { "#\uFE0F\u20E3", ManagedTextShapingTestAssets.EmojiSequenceLigatures[2].Ligature },
        { "\U0001F1F5\U0001F1F1", ManagedTextShapingTestAssets.EmojiSequenceLigatures[3].Ligature },
        { "\U0001F3F4\U000E0067\U000E0062\U000E0065\U000E006E\U000E0067\U000E007F", ManagedTextShapingTestAssets.EmojiSequenceLigatures[4].Ligature }
    };

    [Theory]
    [MemberData(nameof(Sequences))]
    public void SequencePaintsOneLigatureGlyphAndMeasuresItsAdvance(string sequence, int ligature) {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateEmojiSequenceFont(), Family);

        PdfGlyphRun run = font.ShapeText(sequence, PdfTextShapingOptions.ForRendering(font.FontName, PdfTextShapingMode.OpenTypeLigatures));

        Assert.Empty(run.Diagnostics);
        Assert.Equal(sequence, run.ActualText);
        PdfGlyphInfo painted = Assert.Single(run.Glyphs, glyph => glyph.AdvanceWidth1000 != 0);
        Assert.Equal(ligature, painted.GlyphId);
        Assert.All(run.Glyphs.Where(glyph => glyph.GlyphId != ligature), glyph => Assert.Equal(1, glyph.GlyphId));
        Assert.Equal(ManagedTextShapingTestAssets.EmojiAdvance, run.TotalAdvanceWidth1000);
        Assert.Equal(10D, font.MeasureTextWidth(sequence, 10D, PdfTextShapingMode.OpenTypeLigatures), 8);
    }

    [Theory]
    [InlineData(LiftingWeightsMale, false)]
    [InlineData("A " + LiftingWeightsMale + " A", false)]
    [InlineData(LiftingWeightsMale, true)]
    [InlineData("f\u200Di f\u200Ci", false)]
    public void SavedDocumentExtractsTheExactSequence(string text, bool tableCell) {
        PdfCreateOptions options = new PdfCreateOptions { IncludeStandardFontToUnicodeMaps = true }
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(Family, ManagedTextShapingTestAssets.CreateEmojiSequenceFont()));
        PdfDocument document = PdfDocument.Create(doc => doc.Page(page => page.Content(content => {
            PdfTextRun run = PdfTextRun.Normal(text, fontFamily: Family);
            if (tableCell) content.Item(item => item.AddTable(new[] { new[] { PdfTableCell.RichTextCell(new[] { run }) } }, OfficeAlignment.Left));
            else content.AddParagraph(paragraph => paragraph.AddRuns(new[] { run }));
        })), options);

        Assert.Empty(document.AnalyzeTextEncoding());
        Assert.Equal(text, PdfReadDocument.Open(document.ToBytes()).ExtractText().Trim());
    }

    // CBDT and sbix emoji fonts (Noto Color Emoji, Apple Color Emoji) have no glyph outline.
    [Fact]
    public void BitmapOnlyColorGlyphIsReportedInsteadOfWrittenBlank() {
        const string text = "A \U0001F600";
        PdfCreateOptions options = new PdfCreateOptions().RegisterNamedFontFamily(
            new PdfEmbeddedFontFamily("Bitmap emoji", ManagedTextShapingTestAssets.CreateBitmapOnlyColorFont(0x1F600)));
        PdfDocument document = PdfDocument.Create(doc => doc.Page(page => page.Content(content =>
            content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal(text, fontFamily: "Bitmap emoji") })))), options);

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(document.AnalyzeTextEncoding(), item => item.Index == 2);
        Assert.Equal("U+1F600", diagnostic.CodePoint);
        Assert.Contains("bitmap or SVG color image", diagnostic.Message, StringComparison.Ordinal);
        Assert.Throws<PdfTextEncodingPreflightException>(() => document.ToBytes());
    }

    // The painted glyph decides: an outlined ligature over bitmap-only components is paintable,
    // and a bitmap-only ligature over outlined components is not.
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void BitmapOnlyCheckUsesThePaintedLigatureGlyph(bool bitmapLigature, bool reported) {
        int[] bitmapOnly = bitmapLigature
            ? new[] { ManagedTextShapingTestAssets.EmojiSequenceLigatures[0].Ligature }
            : new[] { ManagedTextShapingTestAssets.EmojiGlyph(0x1F3CB), ManagedTextShapingTestAssets.EmojiGlyph(0x2642) };
        PdfCreateOptions options = new PdfCreateOptions().RegisterNamedFontFamily(
            new PdfEmbeddedFontFamily(Family, ManagedTextShapingTestAssets.CreateEmojiSequenceFont(bitmapOnly)));
        PdfDocument document = PdfDocument.Create(doc => doc.Page(page => page.Content(content =>
            content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal(LiftingWeightsMale, fontFamily: Family) })))), options);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = document.AnalyzeTextEncoding();

        if (!reported) {
            Assert.Empty(diagnostics);
            return;
        }
        Assert.Equal(0, Assert.Single(diagnostics).Index);
    }

    // Built-in Latin ligatures hand joiner text to OpenType shaping, which writes the blank
    // space glyph for a ZWJ the font does not map; the preflight checks that same route.
    [Fact]
    public void LatinLigaturePreflightAdmitsJoinerTextThatRenderingShapes() {
        const string text = "A\u200DA";
        PdfCreateOptions options = new PdfCreateOptions { TextShapingMode = PdfTextShapingMode.LatinLigatures }
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Joiner", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true, (int)'A')));
        PdfDocument document = PdfDocument.Create(doc => doc.Page(page => page.Content(content =>
            content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal(text, fontFamily: "Joiner") })))), options);

        Assert.Empty(document.AnalyzeTextEncoding());
        Assert.Equal(text, PdfReadDocument.Open(document.ToBytes()).ExtractText().Trim());
    }

    [Fact]
    public void SegoeUiEmojiZwjSequenceSavesAndExtractsExactly() {
        var options = new PdfCreateOptions { IncludeStandardFontToUnicodeMaps = true };
        // Windows ships the color font; other platforms have no equivalent to resolve.
        if (!options.TryRegisterNamedOfficeFontFamily("Segoe UI Emoji", out string? family)) return;
        const string text = LiftingWeightsMale + " \U0001F468\u200D\U0001F469\u200D\U0001F467\u200D\U0001F466 \u2764\uFE0F #\uFE0F\u20E3";
        PdfDocument document = PdfDocument.Create(doc => doc.Page(page => page.Content(content =>
            content.AddParagraph(paragraph => paragraph.AddRuns(new[] { PdfTextRun.Normal(text, fontFamily: family) })))), options);

        Assert.Equal(text, PdfReadDocument.Open(document.ToBytes()).ExtractText().Trim());
    }
}
