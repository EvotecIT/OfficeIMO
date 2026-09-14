using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgFontScopeTests {
    [Theory]
    [InlineData("\"ACME, Sans\", Arial", "ACME, Sans", "Arial")]
    [InlineData("'ACME, Sans', Arial", "ACME, Sans", "Arial")]
    [InlineData("ACME\\, Sans, Arial", "ACME, Sans", "Arial")]
    public void FontFamilyParserKeepsQuotedAndEscapedCommas(string value, string first, string second) {
        Assert.Equal(new[] { first, second }, OfficeFontFamilyParser.Parse(value));
    }

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

    [Fact]
    public void InlineSvgScopesAQuotedFamilyContainingAComma() {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("ACME, Sans", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30, new OfficeFontInfo("\"ACME, Sans\", Arial", 12));

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(
            drawing, 1, OfficeSvgSizeUnit.Point, null, "page-1-"));

        Assert.Equal("\"page-1-ACME, Sans\", Arial",
            svg.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
    }

    [Fact]
    public void InlineSvgKeepsAnUnmatchedCommaFamilyQuotedWhileScopingTheEmbeddedFallback() {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("ACME, Sans", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30,
            new OfficeFontInfo("\"Missing, Family\", \"ACME, Sans\"", 12));

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(
            drawing, 1, OfficeSvgSizeUnit.Point, null, "page-1-"));

        Assert.Equal("\"Missing, Family\", \"page-1-ACME, Sans\"",
            svg.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
    }

    [Fact]
    public void InlineSvgQuotesAnUnmatchedFamilyThatCannotBeAnUnquotedCssIdentifier() {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("Shared Font", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30,
            new OfficeFontInfo("\"3 of 9 Barcode\", \"Shared Font\"", 12));

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(
            drawing, 1, OfficeSvgSizeUnit.Point, null, "page-1-"));

        Assert.Equal("\"3 of 9 Barcode\", \"page-1-Shared Font\"",
            svg.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
    }

    [Fact]
    public void ScopedSvgRemovesXmlIllegalTextCharactersBeforeReparsing() {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("Shared Font", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A\0\u0001B\ud800C\ufffeD\ud83d\ude00", 2, 2, 90, 30,
            new OfficeFontInfo("Shared Font", 12));

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(
            drawing, 1, OfficeSvgSizeUnit.Point, null, "page-1-"));

        Assert.Equal("ABCD\ud83d\ude00", svg.Descendants().Single(x => x.Name.LocalName == "text").Value);
    }

    private static OfficeDrawing CreateDrawing(char letter) {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("Shared Font", ManagedTextShapingTestAssets.CreateFont(letter));
        return drawing.AddText(letter.ToString(), 2, 2, 90, 30, new OfficeFontInfo("Shared Font", 12));
    }
}
