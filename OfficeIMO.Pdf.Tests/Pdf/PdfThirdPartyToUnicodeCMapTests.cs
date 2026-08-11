using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCompactToUnicodeCMapTests {
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
