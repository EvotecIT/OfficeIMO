using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFontMeasurementTests {
    [Theory]
    [InlineData(false, "office affinity fine flow")]
    [InlineData(true, "office affinity fine flow")]
    [InlineData(false, "123.45 ([)]")]
    [InlineData(true, "123.45 ([)]")]
    [InlineData(false, "A\u00E9 B")]
    [InlineData(true, "A\u00E9\u0301 B")]
    public void DefaultOpenTypeWidthAndSubsetUsageMatchTheRenderedGlyphRun(bool cff, string text) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] data = File.ReadAllBytes(path!);
        var options = PdfTextShapingOptions.ForRendering("Measurement", PdfTextShapingMode.OpenTypeLigatures);
        if (cff) {
            var measured = PdfOpenTypeCffFontProgram.Parse(data, "Measurement");
            var rendered = PdfOpenTypeCffFontProgram.Parse(data, "Measurement");
            PdfGlyphRun run = rendered.ShapeText(text, options);
            Assert.Equal(run.TotalAdvanceWidth1000 * 13.7 / 1000, measured.MeasureTextWidth(text, 13.7, PdfTextShapingMode.OpenTypeLigatures));
            Assert.Equal(rendered.GetGlyphToUnicodeMappings(), measured.GetGlyphToUnicodeMappings());
            Assert.Equal(rendered.GetUsedGlyphIds(), measured.GetUsedGlyphIds());
            if (text.StartsWith("office", StringComparison.Ordinal)) Assert.True(run.Glyphs.Count < text.Length);
        } else {
            var measured = PdfTrueTypeFontProgram.Parse(data, "Measurement");
            var rendered = PdfTrueTypeFontProgram.Parse(data, "Measurement");
            PdfGlyphRun run = rendered.ShapeText(text, options);
            Assert.Equal(run.TotalAdvanceWidth1000 * 13.7 / 1000, measured.MeasureTextWidth(text, 13.7, PdfTextShapingMode.OpenTypeLigatures));
            Assert.Equal(rendered.GetGlyphToUnicodeMappings(), measured.GetGlyphToUnicodeMappings());
            Assert.Equal(rendered.GetUsedGlyphIds(), measured.GetUsedGlyphIds());
            if (text.StartsWith("office", StringComparison.Ordinal)) Assert.True(run.Glyphs.Count < text.Length);
        }
    }

    [Fact]
    public void DefaultMeasurementPreservesSubstitutionBoundariesAndScalarFallback() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn");
        var font = PdfTrueTypeFontProgram.Parse(data, "Measurement");
        Assert.Equal(5, font.MeasureTextWidth("fi", 10, PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(10, font.MeasureTextWidth("f\u200Ei", 10, PdfTextShapingMode.OpenTypeLigatures));
        Assert.Contains(font.GetGlyphToUnicodeMappings(), mapping => mapping.GlyphId == 3 && mapping.UnicodeText == "fi");
        var unsupported = PdfTrueTypeFontProgram.Parse(
            ManagedTextShapingTestAssets.CreateFontWithCommonSubstitution(true), "Unsupported");
        Assert.Equal(15, unsupported.MeasureTextWidth("A,B", 10, PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(10, unsupported.MeasureTextWidth("11", 10, PdfTextShapingMode.OpenTypeLigatures));
        var marked = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithInheritedMarkSubstitution(), "Marked");
        Assert.Equal(15, marked.MeasureTextWidth("A\u0301", 10, PdfTextShapingMode.OpenTypeLigatures));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DefaultMeasurementStillRejectsMissingGlyphs(bool cff) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] data = File.ReadAllBytes(path!);
        const string text = "A\U0001F9D0";
        if (cff) {
            var font = PdfOpenTypeCffFontProgram.Parse(data, "Measurement");
            var expected = Assert.Throws<ArgumentException>(() => font.ShapeText(text, PdfTextShapingOptions.ForRendering("Measurement", PdfTextShapingMode.OpenTypeLigatures)));
            var actual = Assert.Throws<ArgumentException>(() => font.MeasureTextWidth(text, 10, PdfTextShapingMode.OpenTypeLigatures));
            Assert.Equal(expected.Message, actual.Message);
        } else {
            var font = PdfTrueTypeFontProgram.Parse(data, "Measurement");
            var expected = Assert.Throws<ArgumentException>(() => font.ShapeText(text, PdfTextShapingOptions.ForRendering("Measurement", PdfTextShapingMode.OpenTypeLigatures)));
            var actual = Assert.Throws<ArgumentException>(() => font.MeasureTextWidth(text, 10, PdfTextShapingMode.OpenTypeLigatures));
            Assert.Equal(expected.Message, actual.Message);
        }
    }
}
