namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    internal static byte[] CreateFontWithScriptSpecificContinuationKerning() {
        byte[] original = CreateMultipleGsub(scriptTag: "latn", featureTag: "liga");
        var gsub = new byte[82];
        System.Array.Copy(original, gsub, 60);
        System.Array.Copy(original, 60, gsub, 62, 20);
        WriteUInt16(gsub, 4, 62);
        WriteUInt16(gsub, 54, 3);
        WriteUInt16(gsub, 60, 5);
        byte[] gpos = CreatePairPositioningGpos();
        WriteTag(gpos, 12, "latn");
        WriteUInt16(gpos, 72, 4);
        WriteUInt16(gpos, 76, 5);
        return CreateFontFromCmap(CreateFormat12Cmap('A', 1, 32, 2), glyphCount: 6, gsub: gsub, gpos: gpos);
    }

    internal static byte[] CreateFontWithSequencedPairPositioning() {
        var gpos = new byte[102];
        System.Array.Copy(CreatePairPositioningGpos(), gpos, 56);
        const int subtable = 56;
        WriteUInt16(gpos, subtable, 1);
        WriteUInt16(gpos, subtable + 2, 14);
        WriteUInt16(gpos, subtable + 4, 4);
        WriteUInt16(gpos, subtable + 6, 4);
        WriteUInt16(gpos, subtable + 8, 2);
        WriteUInt16(gpos, subtable + 10, 22);
        WriteUInt16(gpos, subtable + 12, 34);
        WriteUInt16(gpos, subtable + 14, 1);
        WriteUInt16(gpos, subtable + 16, 2);
        WriteUInt16(gpos, subtable + 18, 1);
        WriteUInt16(gpos, subtable + 20, 2);
        WriteUInt16(gpos, subtable + 22, 1);
        WriteUInt16(gpos, subtable + 24, 2);
        WriteInt16(gpos, subtable + 26, -10);
        WriteInt16(gpos, subtable + 28, -20);
        WriteUInt16(gpos, subtable + 34, 1);
        WriteUInt16(gpos, subtable + 36, 3);
        WriteInt16(gpos, subtable + 38, -30);
        WriteInt16(gpos, subtable + 40, -40);
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { (int)'A', (int)'B', (int)'C' }), glyphCount: 4, gpos: gpos);
    }
}
