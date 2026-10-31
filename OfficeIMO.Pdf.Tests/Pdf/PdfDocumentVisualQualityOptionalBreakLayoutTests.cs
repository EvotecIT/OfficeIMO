using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OptionalBreaksKeepCharacterJustificationGlyphsAndAdvances(bool embedded) {
        PdfCreateOptions options = JustificationPolicyOptions();
        if (embedded) options.UseFontFamily(new PdfEmbeddedFontFamily("Identifier Font",
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'B', 'C', 'D')));
        // A trailing separator keeps these paragraph tests outside the single-identifier
        // ActualText anchors, so the show operators below describe only painted glyphs.
        byte[] control = RenderJustificationPolicyRuns(options, new[] { new PdfTextRun("ABCD ") }, distributed: true);
        byte[] actual = RenderJustificationPolicyRuns(options, new[] { new PdfTextRun("A\u200BBCD ") }, distributed: true);

        Assert.Equal("ABCD", string.Concat(ReadJustificationPolicyLetters(actual).Select(letter => letter.Value)));
        Assert.Equal(JustificationShownGlyphHex(control), JustificationShownGlyphHex(actual));
        Assert.Equal(JustificationTextAdjustments(control), JustificationTextAdjustments(actual));
        AssertJustificationPolicyWidth(ReadJustificationPolicyLetters(actual).ToList(), expanded: true);
    }

    [Fact]
    public void OptionalBreaksKeepCharacterJustificationCombiningClustersAcrossRuns() {
        PdfCreateOptions options = JustificationPolicyOptions().UseFontFamily(new PdfEmbeddedFontFamily("Cluster Identifier Font",
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'B', 'C', 0x0301)));
        PdfTextRun[] Runs(string prefix) => new[] { new PdfTextRun(prefix), new PdfTextRun("\u0301C ", color: global::OfficeIMO.OfficeColor.FromNormalizedRgb(1, 0, 0)) };
        byte[] control = RenderJustificationPolicyRuns(options, Runs("AB"), distributed: true);
        byte[] actual = RenderJustificationPolicyRuns(options, Runs("A\u200BB"), distributed: true);

        Assert.Equal("AB\u0301C", string.Concat(ReadJustificationPolicyLetters(actual).Select(letter => letter.Value)));
        Assert.Equal(JustificationShownGlyphHex(control), JustificationShownGlyphHex(actual));
        Assert.Equal(JustificationTextAdjustments(control), JustificationTextAdjustments(actual));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OptionalBreaksKeepDecimalTabAtThePaintedSeparator(bool embedded) {
        PdfCreateOptions options = JustificationPolicyOptions();
        if (embedded) options.UseFontFamily(new PdfEmbeddedFontFamily("Numeric Identifier Font",
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'W', '1', '2', '.', '5', '0')));
        var style = new PdfParagraphStyle();
        style.AddTabStop(global::OfficeIMO.OfficeLength.Points(150D), PdfTabAlignment.DecimalSeparator).DecimalSeparator = '.';
        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("\tW\u200B12.50"), style: style).ToBytes();

        using var pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal("W12.50", PdfReadDocument.Open(bytes).ExtractText().Trim());
        var separator = Assert.Single(page.Letters, letter => letter.Value == ".");
        Assert.Equal(180D, separator.StartBaseLine.X, 2);
    }
}
