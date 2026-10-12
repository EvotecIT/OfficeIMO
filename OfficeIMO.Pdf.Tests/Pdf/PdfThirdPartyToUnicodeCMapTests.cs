using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCompactToUnicodeCMapTests {
    [Theory]
    [InlineData("7f ff ff ff", "7FFFFFFF")]
    [InlineData("ff ff ff ff", "FFFFFFFF")]
    public void ToUnicodeCMap_Bounds_single_entry_ranges_at_four_byte_boundaries(string sourceCode, string expectedKey) {
        string source = $"2 beginbfrange <{sourceCode}><{sourceCode}><0041> "
            + "<00 00 00 00><ff ff ff ff><0042> endbfrange";
        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(source), out ToUnicodeCMap? cmap));
        var mapping = Assert.Single(cmap!.Mappings);
        Assert.Equal(expectedKey, mapping.Key);
        Assert.Equal("A", mapping.Value);
        Assert.Equal("A", cmap.MapBytes(Convert.FromHexString(expectedKey)));

        source = $"1 beginbfrange <{sourceCode}><{sourceCode}>[<0041><0042>] endbfrange";
        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(source), out cmap));
        mapping = Assert.Single(cmap!.Mappings);
        Assert.Equal(expectedKey, mapping.Key);
        Assert.Equal("A", mapping.Value);
    }

    [Theory]
    [InlineData("7F FF FF FE", "80 00 00 00", "7FFFFFFE7FFFFFFF80000000")]
    [InlineData("FF FF FF FD", "FF FF FF FF", "FFFFFFFDFFFFFFFEFFFFFFFF")]
    public void ToUnicodeCMap_Retains_unsigned_source_ranges_without_wrapping(string from, string to, string sourceHex) {
        foreach (string destinations in new[] { "<0041>", "[<0041><0042><0043><0044>]" }) {
            string source = $"1 beginbfrange <{from}><{to}>{destinations} endbfrange";
            Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(source), out ToUnicodeCMap? cmap));
            Assert.Equal(3, cmap!.MappingCount);
            Assert.Equal("ABC", cmap.MapBytes(Convert.FromHexString(sourceHex)));
        }
    }

    [Fact]
    public void ToUnicodeCMap_Retains_Apple_surrogate_pairs_and_ranges_with_hex_whitespace() {
        const string source = "1 beginbfchar\n<2 1><d83d de00>\nendbfchar\n"
            + "2 beginbfrange\n<00 22><00 23><00\t41>\n<00\r24><00\n25>[<00\f43><d83d\0de00>]\nendbfrange";

        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(source), out ToUnicodeCMap? cmap));
        Assert.Equal("😀ABC😀", cmap!.MapBytes(new byte[] { 0x21, 0, 0x22, 0, 0x23, 0, 0x24, 0, 0x25 }));
        Assert.True(cmap.TryEncodeText("😀ABC😀", out string encoded));
        Assert.Equal("2100220023002421", encoded);
    }

    [Fact]
    public void ToUnicodeCMap_ParsesAdjacentDelimitedTokensAcrossMappingForms() {
        const string source = """
            begincmap
            1 beginbfchar
            <01><0041>
            endbfchar
            1 beginbfrange
            <02><03>[<0042><0043>]
            endbfrange
            endcmap
            """;

        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(source), out ToUnicodeCMap? cmap));
        Assert.NotNull(cmap);
        Assert.Equal("ABC", cmap.MapBytes(new byte[] { 0x01, 0x02, 0x03 }));
    }

    [Fact]
    public void ToUnicodeCMap_ParsesCompactTwoByteSingleEntryRanges() {
        const string source = """
            /CIDInit /ProcSet findresource begin
            12 dict begin
            begincmap
            /CIDSystemInfo
            << /Registry (Adobe)
            /Ordering (UCS)
            /Supplement 0
            >> def
            /CMapName /Adobe-Identity-UCS def
            /CMapType 2 def
            1 begincodespacerange
            <0000><FFFF>
            endcodespacerange
            3 beginbfrange
            <0002><0002><0020>
            <0006><0006><0042>
            <003b><003b><0065>
            endbfrange
            endcmap
            CMapName currentdict /CMap defineresource pop
            end end
            """;

        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(source), out ToUnicodeCMap? cmap));
        Assert.NotNull(cmap);
        Assert.Equal("B e", cmap.MapBytes(new byte[] { 0x00, 0x06, 0x00, 0x02, 0x00, 0x3B }));
    }

    [Fact]
    public void ToUnicodeCMap_ParsesBomAndCompactDictionary() {
        const string source = """
            ﻿/CIDInit /ProcSet findresource begin
            12 dict begin
            begincmap
            /CIDSystemInfo << /Registry (Adobe)/Ordering (UCS)/Supplement 0>> def
            /CMapName /Adobe-Identity-UCS def /CMapType 2 def
            1 begincodespacerange
            <000F><005C>
            endcodespacerange
            3 beginbfrange
            <0024><0024><0041>
            <0048><0048><0065>
            <0055><0055><0072>
            endbfrange
            endcmap CMapName currentdict /CMap defineresource pop end end
            """;

        Assert.True(ToUnicodeCMap.TryParse(Encoding.UTF8.GetBytes(source), out ToUnicodeCMap? cmap));
        Assert.NotNull(cmap);
        Assert.Equal("Are", cmap.MapBytes(new byte[] { 0x00, 0x24, 0x00, 0x55, 0x00, 0x48 }));
    }
}
