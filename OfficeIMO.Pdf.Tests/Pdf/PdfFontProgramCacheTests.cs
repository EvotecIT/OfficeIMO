using OfficeIMO.Pdf;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFontProgramCacheTests {
    [Fact]
    public void TrueTypeCacheKeepsGlyphUsageIsolatedAcrossDocumentForks() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram first = PdfFontProgramCache.GetTrueType(fontData, "OfficeIMO Cache Test");
        PdfTrueTypeFontProgram second = PdfFontProgramCache.GetTrueType(fontData, "OfficeIMO Cache Test");
        Assert.True(first.TryGetGlyphId('A', out int firstGlyph));
        Assert.True(second.TryGetGlyphId('B', out int secondGlyph));

        first.RecordGlyphUsage(firstGlyph, 'A');
        second.RecordGlyphUsage(secondGlyph, 'B');

        Assert.Equal(new[] { firstGlyph }, first.GetUsedGlyphIds());
        Assert.Equal(new[] { secondGlyph }, second.GetUsedGlyphIds());
    }

    [Fact]
    public void TrueTypeCacheDoesNotReuseBlueprintAfterCallerMutatesAndRepairsBytes() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        byte[] scalerType = fontData.Take(4).ToArray();
        Assert.Empty(PdfFontDiagnostics.AnalyzeEmbeddedFont(fontData, fontName: "Mutable cache test"));

        Array.Clear(fontData, 0, 4);
        Assert.NotEmpty(PdfFontDiagnostics.AnalyzeEmbeddedFont(fontData, fontName: "Mutable cache test"));

        Array.Copy(scalerType, fontData, scalerType.Length);
        Assert.Empty(PdfFontDiagnostics.AnalyzeEmbeddedFont(fontData, fontName: "Mutable cache test"));
    }

    [Fact]
    public void TrueTypeSubsetDropsLayoutAndDeviceMetricTablesButRemainsParseable() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] original = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram program = PdfTrueTypeFontProgram.Parse(original, "Compact subset test");
        Assert.True(program.TryGetGlyphId('A', out int glyphId));
        program.RecordGlyphUsage(glyphId, 'A');

        byte[] subset = program.BuildSubsetFontFile();
        HashSet<string> tables = ReadTrueTypeTableTags(subset);
        PdfTrueTypeFontProgram reparsed = PdfTrueTypeFontProgram.Parse(subset, "Reparsed compact subset");

        Assert.True(subset.Length < original.Length);
        Assert.True(reparsed.TryGetGlyphId('A', out _));
        Assert.DoesNotContain("GDEF", tables);
        Assert.DoesNotContain("GPOS", tables);
        Assert.DoesNotContain("GSUB", tables);
        Assert.DoesNotContain("hdmx", tables);
        Assert.DoesNotContain("vmtx", tables);
        Assert.Contains("glyf", tables);
        Assert.Contains("loca", tables);
    }

    [Fact]
    public void TrueTypeGlyphUsageMapsManagedArabicPresentationFormsToLogicalText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        PdfTrueTypeFontProgram program = PdfTrueTypeFontProgram.Parse(
            File.ReadAllBytes(fontPath),
            "Arabic extraction test");
        const char presentationForm = '\uFE8D';
        if (!program.TryGetGlyphId(presentationForm, out int glyphId) || glyphId <= 0) {
            return;
        }

        program.RecordGlyphUsage(glyphId, presentationForm);

        Assert.Contains(
            program.GetGlyphToUnicodeMappings(),
            mapping => mapping.GlyphId == glyphId && mapping.UnicodeText == "\u0627");
    }

    [Fact]
    public void OpenTypeCffCacheIsThreadSafeAndKeepsDocumentUsageIsolated() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        byte[] fontData = File.ReadAllBytes(fontPath!);
        var programs = new PdfOpenTypeCffFontProgram[32];

        Parallel.For(0, programs.Length, index => {
            PdfOpenTypeCffFontProgram program = PdfFontProgramCache.GetOpenTypeCff(
                fontData,
                "OfficeIMO Cache CFF Test");
            char scalar = (char)('A' + index % 20);
            Assert.True(program.TryGetGlyphId(scalar, out int glyphId));
            program.RecordGlyphUsage(glyphId, scalar);
            programs[index] = program;
        });

        for (int index = 0; index < programs.Length; index++) {
            char scalar = (char)('A' + index % 20);
            Assert.True(programs[index].TryGetGlyphId(scalar, out int expectedGlyph));
            Assert.Equal(new[] { expectedGlyph }, programs[index].GetUsedGlyphIds());
        }
    }

    private static HashSet<string> ReadTrueTypeTableTags(byte[] data) {
        int tableCount = (data[4] << 8) | data[5];
        var tags = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < tableCount; index++) {
            tags.Add(System.Text.Encoding.ASCII.GetString(data, 12 + (index * 16), 4));
        }
        return tags;
    }
}
