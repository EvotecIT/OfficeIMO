using System;
using System.Collections.Generic;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFontSecurityTests {
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
}
