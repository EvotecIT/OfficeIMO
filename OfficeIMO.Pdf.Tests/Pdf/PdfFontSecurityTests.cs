using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFontSecurityTests {
    [Fact]
    public void NamedFontRegistrationRejectsUnboundedDistinctFamilies() {
        var options = new PdfOptions();
        for (int index = 0; index < PdfOptions.MaximumNamedFontFamilies; index++) {
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                "SecurityFont" + index.ToString(), new byte[] { 1 }));
        }

        Assert.Throws<InvalidOperationException>(() =>
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("OverflowFont", new byte[] { 1 })));
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("SecurityFont0", new byte[] { 2 }));
        Assert.Equal(PdfOptions.MaximumNamedFontFamilies, options.NamedFontFamilies.Count);
    }

    [Fact]
    public void ToUnicodeCMap_SkipsOversizedSequentialRanges() {
        byte[] cmapBytes = Encoding.ASCII.GetBytes("""
beginbfchar
<0001> <0041>
endbfchar
beginbfrange
<1000> <FFFF> <0042>
endbfrange
""");

        Assert.True(ToUnicodeCMap.TryParse(cmapBytes, out ToUnicodeCMap? cmap));
        Assert.NotNull(cmap);

        Assert.Equal("A", cmap!.MapBytes(new byte[] { 0x00, 0x01 }));
        Assert.NotEqual("B", cmap.MapBytes(new byte[] { 0x10, 0x00 }));
    }

    [Fact]
    public void ToUnicodeCMap_CapsDuplicateSourceEntriesByProcessedCount() {
        var builder = new StringBuilder();
        builder.AppendLine("beginbfchar");
        for (int index = 0; index < 70000; index++) {
            builder.AppendLine("<01> <0041>");
        }

        builder.AppendLine("endbfchar");

        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(builder.ToString()), out ToUnicodeCMap? cmap));
        Assert.NotNull(cmap);

        FieldInfo field = typeof(ToUnicodeCMap).GetField("_processedMappings", BindingFlags.NonPublic | BindingFlags.Instance)!;
        Assert.Equal(65536, (int)field.GetValue(cmap!)!);
        Assert.Equal("A", cmap!.MapBytes(new byte[] { 0x01 }));
    }

    [Fact]
    public void ToUnicodeCMap_CountsRejectedEntriesTowardMappingBudget() {
        var cmap = new ToUnicodeCMap();
        MethodInfo addMap = typeof(ToUnicodeCMap).GetMethod("AddMap", BindingFlags.NonPublic | BindingFlags.Instance)!;
        FieldInfo processedMappings = typeof(ToUnicodeCMap).GetField("_processedMappings", BindingFlags.NonPublic | BindingFlags.Instance)!;

        addMap.Invoke(cmap, new object[] { "0102030405", "0041" });
        Assert.Equal(1, (int)processedMappings.GetValue(cmap)!);

        addMap.Invoke(cmap, new object[] { "01", "0042" });
        Assert.Equal(2, (int)processedMappings.GetValue(cmap)!);
        Assert.Equal("B", cmap.MapBytes(new byte[] { 0x01 }));
    }

    [Fact]
    public void ToUnicodeCMap_ExcludesOversizedReverseMappings() {
        string repeatedDestination = string.Concat(Enumerable.Repeat("0041", 65));
        byte[] cmapBytes = Encoding.ASCII.GetBytes($"""
beginbfchar
<01> <{repeatedDestination}>
<02> <0041>
endbfchar
""");

        Assert.True(ToUnicodeCMap.TryParse(cmapBytes, out ToUnicodeCMap? cmap));
        Assert.NotNull(cmap);

        Assert.True(cmap!.TryEncodeText(new string('A', 65), out string encoded));
        Assert.Equal(string.Concat(Enumerable.Repeat("02", 65)), encoded);
    }

    [Fact]
    public void ToUnicodeCMap_CapsCumulativeReverseIndexNodes() {
        var cmap = new ToUnicodeCMap();
        MethodInfo addMap = typeof(ToUnicodeCMap).GetMethod("AddMap", BindingFlags.NonPublic | BindingFlags.Instance)!;
        FieldInfo nodeCount = typeof(ToUnicodeCMap).GetField("_reverseMapNodeCount", BindingFlags.NonPublic | BindingFlags.Instance)!;
        FieldInfo budgetExhausted = typeof(ToUnicodeCMap).GetField("_reverseMapBudgetExhausted", BindingFlags.NonPublic | BindingFlags.Instance)!;

        for (int index = 1; index <= 2048; index++) {
            string destination = string.Concat(Enumerable.Repeat(index.ToString("X4"), 64));
            addMap.Invoke(cmap, new object[] { index.ToString("X4"), destination });
        }

        Assert.True((bool)budgetExhausted.GetValue(cmap)!);
        Assert.InRange((int)nodeCount.GetValue(cmap)!, 1, 65536);
    }

    [Fact]
    public void ResourceResolver_CapsCidWidthRangeExpansion() {
        var page = new PdfDictionary();
        var resources = new PdfDictionary();
        var fontDictionary = new PdfDictionary();
        var type0Font = new PdfDictionary();
        var descendant = new PdfDictionary();
        var descendantFonts = new PdfArray();
        var widths = new PdfArray();

        widths.Items.Add(new PdfNumber(0));
        widths.Items.Add(new PdfNumber(100000));
        widths.Items.Add(new PdfNumber(250));
        descendant.Items["DW"] = new PdfNumber(1000);
        descendant.Items["W"] = widths;
        descendantFonts.Items.Add(descendant);
        type0Font.Items["Subtype"] = new PdfName("Type0");
        type0Font.Items["DescendantFonts"] = descendantFonts;
        fontDictionary.Items["F1"] = type0Font;
        resources.Items["Font"] = fontDictionary;
        page.Items["Resources"] = resources;

        Dictionary<string, Func<byte[], double>> providers = ResourceResolver.GetFontWidthProviders(page, new Dictionary<int, PdfIndirectObject>());

        Func<byte[], double> provider = Assert.Contains("F1", providers);
        Assert.Equal(250, provider(new byte[] { 0x00, 0x01 }));
        Assert.Equal(1000, provider(new byte[] { 0x13, 0x87 }));
    }

    [Fact]
    public void OpenTypeInspectorRejectsOversizedFormat12CmapExpansion() {
        byte[] fontData = CreateMinimalTrueTypeFont(CreateLargeRangeFormat12Cmap());

        Assert.False(PdfOpenTypeFontInspector.TryInspect(fontData, out PdfOpenTypeFontInfo? info, out string? error, "OfficeIMO Security Font"));
        Assert.Null(info);
        Assert.Contains("cmap mapping count exceeds supported limits", error, StringComparison.Ordinal);
    }

    [Fact]
    public void TrueTypeFontProgramRejectsOversizedFormat12CmapExpansion() {
        byte[] fontData = CreateMinimalTrueTypeFont(CreateLargeRangeFormat12Cmap());

        NotSupportedException exception = Assert.Throws<NotSupportedException>(
            () => PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMO Security Font"));

        Assert.Contains("cmap mapping count exceeds supported limits", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OpenTypeInspectorRejectsCumulativeOverlappingFormat12CmapExpansion() {
        byte[] fontData = CreateMinimalTrueTypeFont(CreateOverlappingFormat12Cmap());

        Assert.False(PdfOpenTypeFontInspector.TryInspect(fontData, out PdfOpenTypeFontInfo? info, out string? error, "OfficeIMO Security Font"));
        Assert.Null(info);
        Assert.Contains("cmap mapping count exceeds supported limits", error, StringComparison.Ordinal);
    }

    [Fact]
    public void OpenTypeInspectorRejectsCumulativeOverlappingFormat4CmapExpansion() {
        byte[] fontData = CreateMinimalTrueTypeFont(CreateOverlappingFormat4Cmap());

        Assert.False(PdfOpenTypeFontInspector.TryInspect(fontData, out PdfOpenTypeFontInfo? info, out string? error, "OfficeIMO Security Font"));
        Assert.Null(info);
        Assert.Contains("cmap mapping count exceeds supported limits", error, StringComparison.Ordinal);
    }

    private static byte[] CreateLargeRangeFormat12Cmap() {
        var data = new byte[40];
        WriteUInt16(data, 2, 1);
        WriteUInt16(data, 4, 3);
        WriteUInt16(data, 6, 10);
        WriteUInt32(data, 8, 12);
        WriteUInt16(data, 12, 12);
        WriteUInt32(data, 16, 28);
        WriteUInt32(data, 24, 1);
        WriteUInt32(data, 28, 0);
        WriteUInt32(data, 32, 0x10FFFF);
        WriteUInt32(data, 36, 1);
        return data;
    }

    private static byte[] CreateOverlappingFormat12Cmap() {
        const int groupCount = 3;
        var data = new byte[12 + 16 + groupCount * 12];
        WriteUInt16(data, 2, 1);
        WriteUInt16(data, 4, 3);
        WriteUInt16(data, 6, 10);
        WriteUInt32(data, 8, 12);
        WriteUInt16(data, 12, 12);
        WriteUInt32(data, 16, (uint)(16 + groupCount * 12));
        WriteUInt32(data, 24, groupCount);
        for (int group = 0; group < groupCount; group++) {
            int offset = 28 + group * 12;
            WriteUInt32(data, offset, 0);
            WriteUInt32(data, offset + 4, 0xFFFF);
            WriteUInt32(data, offset + 8, 1);
        }

        return data;
    }

    private static byte[] CreateOverlappingFormat4Cmap() {
        const int segmentCount = 4;
        const int subtableLength = 16 + segmentCount * 8;
        var data = new byte[12 + subtableLength];
        WriteUInt16(data, 2, 1);
        WriteUInt16(data, 4, 3);
        WriteUInt16(data, 6, 1);
        WriteUInt32(data, 8, 12);
        WriteUInt16(data, 12, 4);
        WriteUInt16(data, 14, subtableLength);
        WriteUInt16(data, 18, segmentCount * 2);

        int endCodeOffset = 26;
        int startCodeOffset = endCodeOffset + segmentCount * 2 + 2;
        int idDeltaOffset = startCodeOffset + segmentCount * 2;
        int idRangeOffsetOffset = idDeltaOffset + segmentCount * 2;
        for (int segment = 0; segment < segmentCount - 1; segment++) {
            WriteUInt16(data, endCodeOffset + segment * 2, 0xFFFE);
            WriteUInt16(data, startCodeOffset + segment * 2, 0);
            WriteUInt16(data, idDeltaOffset + segment * 2, 1);
            WriteUInt16(data, idRangeOffsetOffset + segment * 2, 0);
        }

        WriteUInt16(data, endCodeOffset + (segmentCount - 1) * 2, 0xFFFF);
        WriteUInt16(data, startCodeOffset + (segmentCount - 1) * 2, 0xFFFF);
        WriteUInt16(data, idDeltaOffset + (segmentCount - 1) * 2, 1);
        WriteUInt16(data, idRangeOffsetOffset + (segmentCount - 1) * 2, 0);
        return data;
    }

    private static byte[] CreateMinimalTrueTypeFont(byte[] cmap) {
        var tables = new List<(string Tag, byte[] Data)> {
            ("cmap", cmap),
            ("glyf", new byte[4]),
            ("head", CreateHeadTable()),
            ("hhea", CreateHheaTable()),
            ("hmtx", CreateHmtxTable()),
            ("maxp", new byte[] { 0x00, 0x01, 0x00, 0x00, 0x00, 0x02 }),
            ("name", new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x06 })
        };

        int tableDirectoryLength = 12 + tables.Count * 16;
        var offsets = new int[tables.Count];
        int offset = tableDirectoryLength;
        for (int index = 0; index < tables.Count; index++) {
            offsets[index] = offset;
            offset += Align4(tables[index].Data.Length);
        }

        var font = new byte[offset];
        WriteUInt32(font, 0, 0x00010000);
        WriteUInt16(font, 4, (ushort)tables.Count);
        for (int index = 0; index < tables.Count; index++) {
            int record = 12 + index * 16;
            WriteTag(font, record, tables[index].Tag);
            WriteUInt32(font, record + 8, (uint)offsets[index]);
            WriteUInt32(font, record + 12, (uint)tables[index].Data.Length);
            Array.Copy(tables[index].Data, 0, font, offsets[index], tables[index].Data.Length);
        }

        return font;
    }

    private static byte[] CreateHeadTable() {
        var table = new byte[54];
        WriteUInt16(table, 18, 1000);
        return table;
    }

    private static byte[] CreateHheaTable() {
        var table = new byte[36];
        WriteUInt16(table, 4, 800);
        WriteUInt16(table, 6, unchecked((ushort)-200));
        WriteUInt16(table, 34, 2);
        return table;
    }

    private static byte[] CreateHmtxTable() {
        var table = new byte[8];
        WriteUInt16(table, 0, 500);
        WriteUInt16(table, 4, 500);
        return table;
    }

    private static int Align4(int value) => (value + 3) & ~3;

    private static void WriteTag(byte[] data, int offset, string tag) {
        for (int index = 0; index < 4; index++) {
            data[offset + index] = (byte)tag[index];
        }
    }

    private static void WriteUInt16(byte[] data, int offset, int value) {
        data[offset] = (byte)((value >> 8) & 0xFF);
        data[offset + 1] = (byte)(value & 0xFF);
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)((value >> 24) & 0xFF);
        data[offset + 1] = (byte)((value >> 16) & 0xFF);
        data[offset + 2] = (byte)((value >> 8) & 0xFF);
        data[offset + 3] = (byte)(value & 0xFF);
    }
}
