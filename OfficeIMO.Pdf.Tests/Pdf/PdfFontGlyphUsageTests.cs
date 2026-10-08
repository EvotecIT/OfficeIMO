using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFontGlyphUsageTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnicodeAliasesAndEmptyContinuationsRetainSubsetMembershipAcrossReset(bool cff) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] data = File.ReadAllBytes(path!);
        PdfTrueTypeFontProgram? tt = cff ? null : PdfTrueTypeFontProgram.Parse(data);
        PdfOpenTypeCffFontProgram? ot = cff ? PdfOpenTypeCffFontProgram.Parse(data) : null;
        int glyphId;
        Assert.True(cff ? ot!.TryGetGlyphId('A', out glyphId) : tt!.TryGetGlyphId('A', out glyphId));
        Action<int> recordScalar = scalar => { if (cff) ot!.RecordGlyphUsage(glyphId, scalar); else tt!.RecordGlyphUsage(glyphId, scalar); };
        Action<string> recordText = text => { if (cff) ot!.RecordGlyphUsage(glyphId, text); else tt!.RecordGlyphUsage(glyphId, text); };
        Func<IReadOnlyList<int>> used = cff ? ot!.GetUsedGlyphIds : tt!.GetUsedGlyphIds;
        Func<IReadOnlyList<(int GlyphId, string UnicodeText)>> mappings = cff ? ot!.GetGlyphToUnicodeMappings : tt!.GetGlyphToUnicodeMappings;

        recordScalar('A');
        recordText("fi");
        recordScalar('B');
        recordText("ffi");
        recordText("aba");
        recordText(string.Empty);
        Assert.Equal(new[] { glyphId }, used());
        Assert.Equal((glyphId, "aba"), Assert.Single(mappings()));

        if (cff) ot!.ResetGlyphUsage(); else tt!.ResetGlyphUsage();
        recordText(string.Empty);
        Assert.Equal(new[] { glyphId }, used());
        recordScalar('C');
        Assert.Equal(new[] { glyphId }, used());
        Assert.Equal((glyphId, "C"), Assert.Single(mappings()));
    }
}
