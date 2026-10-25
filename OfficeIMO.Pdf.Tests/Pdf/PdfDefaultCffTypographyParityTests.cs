using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDefaultCffTypographyParityTests {
    [Theory]
    [InlineData("OfficeIMO.Pdf.Tests/Pdf/Fixtures/Fonts/SourceSerif4-Regular.otf", "\uFB01 office e\u0301")]
    [InlineData("OfficeIMO.Drawing.Tests/TestAssets/SourceSansPro-Regular.otf", "a\u0323\u0301 office")]
    [InlineData("OfficeIMO.Pdf.Tests/Pdf/Fixtures/Fonts/SourceSerif4-Regular.otf", "office AVATAR")]
    public void DefaultCffTypographyAndMeasurementRetainCanonicalGlyphsAndClusters(string relativePath, string text) {
        byte[] data = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), relativePath));
        var font = PdfOpenTypeCffFontProgram.Parse(data, "CffTypography");
        IOfficeFontProgram core = Assert.IsType<OfficeOpenTypeCffFont>(OfficeOpenTypeCffFont.TryLoad(data, null, out _));
        var request = new OfficeTextShapingRequest(text, "CffTypography", data, true, font.UnitsPerEm,
            OfficeTextDirection.Auto, null, default, fontCollectionIndex: null, variationCoordinates: null,
            cloneFontData: true, applyDefaultLatinLigatures: true);
        OfficeTextShapingResult expected = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(request));
        PdfTextShapingOptions defaults = PdfTextShapingOptions.ForRendering("CffTypography", PdfTextShapingMode.OpenTypeLigatures);
        PdfTextShapingOptions explicitManaged = PdfTextShapingOptions.ForRendering("CffTypography", PdfTextShapingMode.OpenTypeLigatures,
            OfficeManagedTextShapingProvider.Instance);
        PdfGlyphRun actual = font.ShapeText(text, defaults);
        PdfGlyphRun explicitRun = font.ShapeText(text, explicitManaged);

        Assert.Equal(expected.Direction, actual.Direction);
        Assert.Equal(explicitRun.Glyphs, actual.Glyphs);
        Assert.Equal(expected.Glyphs.Count, actual.Glyphs.Count);
        for (int index = 0; index < actual.Glyphs.Count; index++) {
            OfficeShapedGlyph shaped = expected.Glyphs[index];
            PdfGlyphInfo glyph = actual.Glyphs[index];
            Assert.Equal(shaped.GlyphId, glyph.GlyphId);
            Assert.Equal(shaped.UnicodeText, glyph.UnicodeText);
            Assert.Equal(shaped.TextIndex, glyph.TextIndex);
            Assert.Equal(expected.GetLogicalClusterStart(index), glyph.LogicalClusterStart);
            Assert.Equal(font.GetGlyphWidth1000(shaped.GlyphId), glyph.NominalWidth1000);
            Assert.Equal((shaped.AdvanceWidth.HasValue ? Scale(shaped.AdvanceWidth.Value, font.UnitsPerEm) : glyph.NominalWidth1000)
                + Scale(expected.GetAdvanceAdjustment(index), font.UnitsPerEm), glyph.AdvanceWidth1000);
            Assert.Equal(Scale(shaped.OffsetX, font.UnitsPerEm), glyph.OffsetX1000);
            Assert.Equal(Scale(shaped.OffsetY, font.UnitsPerEm), glyph.OffsetY1000);
        }
        Assert.Equal(text, string.Concat(actual.Glyphs.OrderBy(glyph => glyph.TextIndex).Select(glyph => glyph.UnicodeText)));
        bool complex = OfficeTextElements.ContainsCombiningMark(text);
        Assert.Equal(complex ? text : null, actual.ActualText);
        foreach (double size in new[] { 10D, 20D }) {
            Assert.Equal(actual.TotalAdvanceWidth1000 * size / 1000D,
                font.MeasureTextWidth(text, size, PdfTextShapingMode.OpenTypeLigatures), 10);
            Assert.Equal(font.MeasureTextWidth(text, size, PdfTextShapingMode.OpenTypeLigatures,
                    OfficeManagedTextShapingProvider.Instance),
                font.MeasureTextWidth(text, size, PdfTextShapingMode.OpenTypeLigatures), 10);
            Assert.Equal(core.MeasureShapedText(text, expected, size),
                font.MeasureTextWidth(text, size, PdfTextShapingMode.OpenTypeLigatures, useFontUnitAdvances: true), 10);
            Assert.Equal(font.MeasureTextWidth(text, size, PdfTextShapingMode.OpenTypeLigatures,
                    OfficeManagedTextShapingProvider.Instance, useFontUnitAdvances: true),
                font.MeasureTextWidth(text, size, PdfTextShapingMode.OpenTypeLigatures, useFontUnitAdvances: true), 10);
        }
        var mappings = font.GetGlyphToUnicodeMappings();
        int[] used = font.GetUsedGlyphIds().ToArray();
        font.ResetGlyphUsage();
        PdfGlyphRun replay = font.ShapeText(text, defaults);
        Assert.Equal(actual.Glyphs, replay.Glyphs);
        Assert.Equal(mappings, font.GetGlyphToUnicodeMappings());
        Assert.Equal(used, font.GetUsedGlyphIds());
    }

    private static int Scale(int value, int unitsPerEm) =>
        (int)Math.Round(value * 1000D / unitsPerEm, MidpointRounding.AwayFromZero);
}
