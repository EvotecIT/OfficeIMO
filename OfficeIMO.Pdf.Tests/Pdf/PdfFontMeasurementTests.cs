using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFontMeasurementTests {
    [Theory]
    [InlineData(false, 16, 8000, 4095938)]
    [InlineData(true, 16, 8000, 4095938)]
    [InlineData(false, 2048, 63, 32000)]
    [InlineData(true, 2048, 63, 32000)]
    public void NominalWidthsPreserveRoundingAndRepeatedHorizontalMetricsAcrossDocumentForks(bool cff, int unitsPerEm, int secondWidth, int lastWidth) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] data = File.ReadAllBytes(path!);
        WriteUInt16(data, FindTableOffset(data, "head") + 18, unitsPerEm);
        WriteUInt16(data, FindTableOffset(data, "hhea") + 34, 3);
        int metrics = FindTableOffset(data, "hmtx");
        WriteUInt16(data, metrics, 0);
        WriteUInt16(data, metrics + 4, 128);
        WriteUInt16(data, metrics + 8, ushort.MaxValue);

        int glyphCount;
        Func<int, int> width;
        Func<int, int> forkWidth;
        PdfTrueTypeFontProgram? trueType = null;
        if (cff) {
            var font = PdfOpenTypeCffFontProgram.Parse(data, "Metric rounding");
            glyphCount = font.GlyphCount;
            width = font.GetGlyphWidth1000;
            forkWidth = font.ForkForDocument().GetGlyphWidth1000;
        } else {
            trueType = PdfTrueTypeFontProgram.Parse(data, "Metric rounding");
            glyphCount = trueType.GlyphCount;
            width = trueType.GetGlyphWidth1000;
            forkWidth = trueType.ForkForDocument().GetGlyphWidth1000;
        }

        Assert.True(glyphCount > 3);
        for (int glyph = 0; glyph < glyphCount; glyph++) {
            int expected = glyph == 0 ? 0 : glyph == 1 ? secondWidth : lastWidth;
            Assert.Equal(expected, width(glyph));
            Assert.Equal(expected, forkWidth(glyph));
        }
        foreach (int glyph in new[] { int.MinValue, -1, glyphCount, int.MaxValue }) {
            Assert.Equal(lastWidth, width(glyph));
            Assert.Equal(lastWidth, forkWidth(glyph));
        }

        if (trueType != null) {
            int[] winAnsiWidths = trueType.BuildWinAnsiWidths();
            for (int code = 32; code <= 255; code++) {
                char character = PdfWinAnsiEncoding.Decode((byte)code);
                int expected = trueType.TryGetGlyphId(character, out int glyph) ? width(glyph) : 500;
                Assert.Equal(expected, winAnsiWidths[code - 32]);
            }
        }
    }

    private static int FindTableOffset(byte[] data, string tag) {
        int count = data[4] * 256 + data[5];
        for (int index = 0; index < count; index++) {
            int entry = 12 + index * 16;
            if (Encoding.ASCII.GetString(data, entry, 4) == tag) {
                return (data[entry + 8] << 24) | (data[entry + 9] << 16) | (data[entry + 10] << 8) | data[entry + 11];
            }
        }
        throw new InvalidOperationException("The metric fixture has no " + tag + " table.");
    }

    private static void WriteUInt16(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

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
