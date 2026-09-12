using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgFontScopeTests {
    [Fact]
    public void InlinePagesKeepSameNamedFontsIndependent() {
        var first = CreateDrawing('A');
        var second = CreateDrawing('B');
        string firstSvg = OfficeDrawingSvgExporter.ToSvg(first, 1, OfficeSvgSizeUnit.Point, null, "page-1-");
        string secondSvg = OfficeDrawingSvgExporter.ToSvg(second, 1, OfficeSvgSizeUnit.Point, null, "page-2-");
        var firstXml = XElement.Parse(firstSvg);
        var secondXml = XElement.Parse(secondSvg);
        Assert.Contains("font-family:\"page-1-Shared Font\"", firstXml.Descendants().Single(x => x.Name.LocalName == "style").Value);
        Assert.Contains("font-family:\"page-2-Shared Font\"", secondXml.Descendants().Single(x => x.Name.LocalName == "style").Value);
        Assert.Equal("\"page-1-Shared Font\"", firstXml.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
        Assert.Equal("\"page-2-Shared Font\"", secondXml.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
        Assert.Equal("A", firstXml.Descendants().Single(x => x.Name.LocalName == "text").Value);
        Assert.Equal("B", secondXml.Descendants().Single(x => x.Name.LocalName == "text").Value);
        Assert.Equal(firstSvg, OfficeDrawingSvgExporter.ToSvg(first, 1, OfficeSvgSizeUnit.Point, null, "page-1-"));
        Assert.Equal("Shared Font", first.Fonts.Faces[0].FamilyName);
    }

    private static OfficeDrawing CreateDrawing(char letter) {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("Shared Font", ManagedTextShapingTestAssets.CreateFont(letter));
        return drawing.AddText(letter.ToString(), 2, 2, 90, 30, new OfficeFontInfo("Shared Font", 12));
    }
}
