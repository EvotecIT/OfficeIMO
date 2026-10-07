using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

internal static class CffRandomGlyphTestAssets {
    internal static byte[] CreateRandomOverhangFont() {
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeCffFontData cff = OfficeCffFontData.Parse(reader, OfficeFontVariationModel.None);
        OfficeCffFontData.CffSlice glyph = cff.GetCharString(reader.MapGlyph('A'));
        // Keep the registered font's table and index layout while replacing one
        // glyph with a supported random-dependent triangle and negative overhang.
        byte[] program = {
            12, 23, 28, 0xF0, 0x60, 12, 24, 28, 0, 0, 21, // random * -4000, 0 rmoveto
            239, 139, 139, 239, 39, 39, 5, 14 // triangle, endchar
        };
        Assert.Equal(glyph.Length, program.Length);
        Array.Copy(program, 0, glyph.Data, glyph.Offset, program.Length);
        return glyph.Data;
    }
}
