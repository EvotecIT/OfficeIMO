namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    /// <summary>Creates independent glyphs for all four Arabic joining forms.</summary>
    internal static byte[] CreateArabicJoiningFont() {
        const int features = 36;
        const int lookups = 86;
        var gsub = new byte[184];
        WriteUInt32(gsub, 0, 0x00010000);
        WriteUInt16(gsub, 4, 10);
        WriteUInt16(gsub, 6, features);
        WriteUInt16(gsub, 8, lookups);
        WriteUInt16(gsub, 10, 1);
        WriteTag(gsub, 12, "arab");
        WriteUInt16(gsub, 16, 8);
        WriteUInt16(gsub, 18, 4);
        WriteUInt16(gsub, 24, 0xFFFF);
        WriteUInt16(gsub, 26, 4);
        WriteUInt16(gsub, features, 4);
        WriteUInt16(gsub, lookups, 4);
        string[] tags = { "fina", "init", "isol", "medi" };
        ushort[] replacements = { 5, 4, 7, 6 };
        for (int index = 0; index < 4; index++) {
            WriteUInt16(gsub, 28 + index * 2, (ushort)index);
            int record = features + 2 + index * 6;
            int feature = 62 + index * 6;
            WriteTag(gsub, record, tags[index]);
            WriteUInt16(gsub, record + 4, (ushort)(feature - features));
            WriteUInt16(gsub, feature + 2, 1);
            WriteUInt16(gsub, feature + 4, (ushort)index);
            int lookup = 96 + index * 22;
            WriteUInt16(gsub, lookups + 2 + index * 2, (ushort)(lookup - lookups));
            WriteUInt16(gsub, lookup, 1);
            WriteUInt16(gsub, lookup + 4, 1);
            WriteUInt16(gsub, lookup + 6, 8);
            int subtable = lookup + 8;
            WriteUInt16(gsub, subtable, 2);
            WriteUInt16(gsub, subtable + 2, 8);
            WriteUInt16(gsub, subtable + 4, 1);
            WriteUInt16(gsub, subtable + 6, replacements[index]);
            WriteUInt16(gsub, subtable + 8, 1);
            WriteUInt16(gsub, subtable + 10, 1);
            WriteUInt16(gsub, subtable + 12, 1);
        }
        return CreateFontFromCmap(CreateFormat12Cmap(0x0628, 1, 32, 2), glyphCount: 8, gsub: gsub);
    }
}
