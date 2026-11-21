using System.Text;
using OfficeIMO.Drawing;
using Xunit;
using OfficeIMO.TestAssets;

namespace OfficeIMO.Drawing.HarfBuzz.Tests;

/// <summary>
/// Compares the managed OpenType engine with HarfBuzz glyph IDs, clusters and pen positions over
/// mark-heavy runs: repository fonts for real lookup data, and generated fonts for lookup flags,
/// context formats and cursive attachment that the repository fonts do not exercise.
/// </summary>
public sealed class ManagedShapingParityTests {
    public static IEnumerable<object[]> RepositoryFontCases() {
        string[] arabic = {
            "بِسْمِ ٱللَّهِ ٱلرَّحْمَٰنِ ٱلرَّحِيمِ", "مُحَمَّدٌ", "لِّلَّهِ", "لَّا", "لَا", "سَلَامٌ", "فِيْ", "شَدَّةٌ", "قُرْآنٌ", "إِنَّ", "لُغَةٌ",
            "الْعَرَبِيَّةُ", "لِلَّه", "بٌّ", "هٰذَا", "تَّ", "لاَ", "لأَ", "لآ", "فِلَسْطِينُ", "اللّٰهُ", "يَسْتَطِيعُونَ"
        };
        foreach (string text in arabic) yield return new object[] { "NotoSansArabic-Regular.ttf", text };
        string[] latin = {
            "a\u0301\u0323", "e\u0302\u0301", "o\u0308\u0304", "Q\u0303", "fi\u0323", "ffi\u0301", "fl\u0300", "A\u030A\u0301", "i\u0307\u0301",
            "g\u0327", "s\u0326\u0301", "\u1E68\u0301", "x\u0323\u0302\u0301", "office\u0301", "n\u0303\u0303\u0303"
        };
        foreach (string font in new[] { "SourceSansPro-Regular.otf", "SourceSerif4-Regular.otf", "OpenSans-Bold.ttf" })
            foreach (string text in latin) yield return new object[] { font, text };
        // Carlito has no combining-mark glyphs: these runs compose to letters it covers.
        foreach (string text in new[] { "A\u030A\u0301", "o\u0308\u0304", "e\u0302\u0301", "fi\u0323", "ffi\u0301", "office\u0301", "g\u0327" })
            yield return new object[] { "Carlito-Regular.ttf", text };
    }

    [Theory]
    [MemberData(nameof(RepositoryFontCases))]
    public void MatchesHarfBuzzOnRepositoryFonts(string fontFile, string text) =>
        AssertParity(File.ReadAllBytes(FontPath(fontFile)), text);

    [Theory]
    [InlineData(false, "abc")]
    [InlineData(false, "ab\u0301c")]
    [InlineData(false, "a\u0301bc\u0301")]
    [InlineData(false, "abc\u0301abc")]
    [InlineData(false, "\u05D0\u05D1\u05D2")]
    [InlineData(false, "\u05D0\u05D1\u05B4\u05D2")]
    [InlineData(true, "abc")]
    [InlineData(true, "ab\u0301c")]
    [InlineData(true, "a\u0301bc\u0301a")]
    [InlineData(true, "\u05D0\u05D1\u05D2")]
    [InlineData(true, "\u05D0\u05B4\u05D1\u05D2\u05B4")]
    public void MatchesHarfBuzzCursiveAttachment(bool rightToLeftFlag, string text) =>
        AssertParity(ManagedTextShapingTestAssets.CreateCursiveAttachmentFont(rightToLeftFlag), text);

    [Theory]
    [InlineData("ab\u0301c")]
    [InlineData("a\u0301bc")]
    [InlineData("c\u0301a")]
    [InlineData("ca\u0301")]
    [InlineData("a\u0301x")]
    [InlineData("ax\u0301")]
    [InlineData("abca\u0301x")]
    [InlineData("cac\u0301a\u0301")]
    public void MatchesHarfBuzzContextFormatsAndNestedFlags(string text) =>
        AssertParity(ManagedTextShapingTestAssets.CreateContextFormatsFont(), text);

    private static void AssertParity(byte[] fontData, string text) {
        OfficeFontFace face = Assert.Single(new OfficeFontFaceCollection().Add("Parity", fontData).Faces);
        bool rightToLeft = text.Any(character => character is >= '\u0590' and <= '\u08FF');
        string language = text.Any(character => character is >= '\u0600' and <= '\u08FF') ? "ar" : rightToLeft ? "he" : "en";
        var request = new OfficeTextShapingRequest(text, "Parity", face.Program.GetFontDataForShaping(), face.Program.IsOpenTypeCff,
            face.Program.UnitsPerEm, rightToLeft ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight, language);
        OfficeTextShapingResult expected = Assert.IsType<OfficeTextShapingResult>(OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(request));
        OfficeTextShapingResult? actual = OfficeManagedTextShapingProvider.Instance.ShapeText(request);
        Assert.True(actual != null, "The managed provider declined " + Escape(text) + ".");
        Assert.Equal(Describe(expected, fontData), Describe(actual!, fontData));
    }

    private static string Escape(string text) {
        var builder = new StringBuilder();
        foreach (char character in text) builder.Append(character < 128 ? character.ToString() : $"\\u{(int)character:X4}");
        return builder.ToString();
    }

    internal static string Describe(OfficeTextShapingResult result, byte[] fontData) {
        int[] advances = ReadAdvances(fontData);
        var builder = new StringBuilder();
        int pen = 0;
        for (int index = 0; index < result.Glyphs.Count; index++) {
            OfficeShapedGlyph glyph = result.Glyphs[index];
            int advance = (glyph.AdvanceWidth ?? advances[Math.Min(glyph.GlyphId, advances.Length - 1)]) + result.GetAdvanceAdjustment(index);
            builder.Append($"{glyph.GlyphId}@{result.GetLogicalClusterStart(index)}({pen + glyph.OffsetX},{glyph.OffsetY}) ");
            pen += advance;
        }
        builder.Append("w=").Append(pen);
        return builder.ToString();
    }

    private static int[] ReadAdvances(byte[] data) {
        int tables = U16(data, 4), hhea = 0, hmtx = 0, maxp = 0;
        for (int index = 0; index < tables; index++) {
            int record = 12 + index * 16;
            string tag = Encoding.ASCII.GetString(data, record, 4);
            int offset = (int)U32(data, record + 8);
            if (tag == "hhea") hhea = offset; else if (tag == "hmtx") hmtx = offset; else if (tag == "maxp") maxp = offset;
        }
        int metrics = U16(data, hhea + 34), glyphs = U16(data, maxp + 4);
        var result = new int[glyphs];
        for (int glyph = 0; glyph < glyphs; glyph++)
            result[glyph] = U16(data, hmtx + Math.Min(glyph, metrics - 1) * 4);
        return result;
    }

    private static int U16(byte[] data, int offset) => (data[offset] << 8) | data[offset + 1];
    private static uint U32(byte[] data, int offset) => ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];

    private static string FontPath(string fileName) => Path.Combine(AppContext.BaseDirectory, "Fonts", fileName);
}
