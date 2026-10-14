using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class VisioRichTextBulletTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void ProducerRoundBulletsRetainNativeRowsAndTextAcrossReopening(int mode) {
        var document = Reopen(VisioDocument.LoadLegacyXml(Path.Combine(AppContext.BaseDirectory,
            "Fixtures", "LegacyXml", "shorewall-netfilter.vdx")).Value, mode);
        byte[] before = document.ToLegacyXmlResult().Value;
        var page = document.Pages[0];
        var shape = page.Shapes.Single(s => s.Id == "32");
        var label = XDocument.Parse(page.ToSvg()).Descendants(Svg + "g")
            .Single(t => (string?)t.Attribute("data-visio-shape-id") == "32");
        var text = label.Descendants(Svg + "text").Where(t => !string.IsNullOrWhiteSpace(t.Value)).ToArray();
        Assert.Equal(new[] { "•", "Raw", "•", "Mangle", "•", "Nat" }, text.Select(t => t.Value));
        for (int i = 0; i < text.Length; i += 2) {
            Assert.InRange(Number(text[i + 1], "x") - Number(text[i], "x"), 23.99, 24.01);
            Assert.Equal(text[i].Attribute("font-size")!.Value, text[i + 1].Attribute("font-size")!.Value);
        }
        Assert.Equal("Raw\nMangle\nNat\n", shape.Text);
        _ = page.ToPng(new VisioPngSaveOptions { Supersampling = 1 });
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }
}
