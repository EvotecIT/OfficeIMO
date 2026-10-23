using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNormalizedEncodingCoverageTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GeneratedEncodingUsesTheCompleteNormalizedRunWithoutChangingOriginalSource(bool named) {
        const string source = "A\u0301";
        const string family = "Normalized encoding";
        // The original combining acute has no cmap entry. NFC composition is
        // sufficient: this is not a new attachment or fallback-font fixture.
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true, 'A', 0x00C1);
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add(family, data).Faces).Program;
        Assert.True(font.TryGetGlyphMetrics('A', out _, out _));
        Assert.True(font.TryGetGlyphMetrics(0x00C1, out int compositeGlyph, out _));
        Assert.False(font.TryGetGlyphMetrics(0x0301, out _, out _));
        OfficeTextShapingResult complete = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(new OfficeTextShapingRequest(
                source, family, data, isOpenTypeCff: false, font.UnitsPerEm,
                OfficeTextDirection.Auto, "en", default, fontCollectionIndex: null,
                variationCoordinates: null, cloneFontData: true, applyDefaultLatinLigatures: true)));
        OfficeShapedGlyph glyph = Assert.Single(complete.Glyphs);
        Assert.Equal(compositeGlyph, glyph.GlyphId);
        Assert.Equal(source, glyph.UnicodeText);
        Assert.Equal(0, glyph.TextIndex);
        Assert.Equal(0, complete.GetLogicalClusterStart(0));
        Assert.Equal(500, glyph.AdvanceWidth);

        var report = new PdfConversionReport();
        PdfOptions options = (named
            ? new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family, data))
            : new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, data, family))
            .SetTextShapingMode(PdfTextShapingMode.OpenTypeLigatures)
            .ReportDiagnosticsTo(report);
        PdfDocument document = PdfDocument.Create(options).Paragraph(paragraph => {
            if (named) paragraph.FontFamily(family);
            paragraph.Text(source);
        });

        Assert.Empty(document.AnalyzeTextEncoding());
        byte[] bytes = document.ToBytes();
        Assert.NotEmpty(bytes);
        // This single source cluster has no spatial-order ambiguity. Exact NFD
        // text proves the painted NFC glyph did not rewrite or duplicate source.
        Assert.Equal(source, PdfReadDocument.Open(bytes).ExtractText().Trim());
        report.RequireNoLoss();
    }

    [Fact]
    public void ExplicitScalarEncodingStillRejectsAnUnmappedOriginalMark() {
        const string source = "A\u0301";
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true, 'A', 0x00C1);
        var options = new PdfOptions()
            .EmbedStandardFont(PdfStandardFont.Helvetica, data, "Scalar encoding")
            .SetTextShapingMode(PdfTextShapingMode.UnicodeScalar);
        PdfDocument document = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text(source));

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(document.AnalyzeTextEncoding());
        Assert.Equal(1, diagnostic.Index);
        Assert.Equal("U+0301", diagnostic.CodePoint);
        Assert.False(diagnostic.IsControlCharacter);
        System.ArgumentException error = Assert.ThrowsAny<System.ArgumentException>(() => document.ToBytes());
        Assert.Equal("U+0301", error.Data["codePoint"]);
    }

}
