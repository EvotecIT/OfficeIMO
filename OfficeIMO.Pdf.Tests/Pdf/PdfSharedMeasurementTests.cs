using OfficeIMO.Pdf;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSharedMeasurementTests {
    [Fact]
    public void WarmWidthsReplayOnlyRequestedGlyphUsageIntoEachFontFork() {
        string? path = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] data = File.ReadAllBytes(path!);
        PdfTrueTypeFontProgram warm = PdfFontProgramCache.GetTrueType(data, "Shared width test");
        double expected = warm.MeasureTextWidth("ABC", 12D, PdfTextShapingMode.OpenTypeLigatures);
        _ = warm.MeasureTextWidth("XYZ", 12D, PdfTextShapingMode.OpenTypeLigatures);
        PdfTrueTypeFontProgram first = PdfFontProgramCache.GetTrueType(data, "Shared width test");
        PdfTrueTypeFontProgram second = PdfFontProgramCache.GetTrueType(data, "Shared width test");

        Parallel.Invoke(
            () => Assert.Equal(expected, first.MeasureTextWidth("ABC", 12D, PdfTextShapingMode.OpenTypeLigatures)),
            () => Assert.Equal(expected, second.MeasureTextWidth("ABC", 12D, PdfTextShapingMode.OpenTypeLigatures)));

        Assert.Equal(new[] { "A", "B", "C" }, first.GetGlyphToUnicodeMappings().Select(p => p.UnicodeText).OrderBy(p => p));
        Assert.Equal(first.GetUsedGlyphIds(), second.GetUsedGlyphIds());
        Assert.Equal(first.BuildSubsetFontFile(), second.BuildSubsetFontFile());
    }
}
