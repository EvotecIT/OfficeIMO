using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgFontScopeTests {
    [Theory]
    [InlineData("ABCDEF+Subset-Regular", "\"ABCDEF+Subset-Regular\"")]
    [InlineData("3 of 9 Barcode", "\"3 of 9 Barcode\"")]
    [InlineData("initial", "\"initial\"")]
    [InlineData("Shared Font", "Shared Font")]
    [InlineData("Shared  Font", "\"Shared  Font\"")]
    public void StandaloneSvgUsesCssSafeEmbeddedFontReferences(string family, string expectedReference) {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add(family, ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30, new OfficeFontInfo(family, 12));

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));

        Assert.Equal(expectedReference,
            svg.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
        Assert.Contains("font-family:\"" + family + "\"",
            svg.Descendants().Single(x => x.Name.LocalName == "style").Value);
        Assert.Equal(family, drawing.Fonts.Faces[0].FamilyName);
    }

    [Fact]
    public void StandaloneSvgPreservesQuotedFallbackNamesAndGenericFamilies() {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("ABCDEF+Subset-Regular", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30,
            new OfficeFontInfo("\"Missing, Family\", ABCDEF+Subset-Regular, sans-serif", 12));

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));

        Assert.Equal("\"Missing, Family\", \"ABCDEF+Subset-Regular\", sans-serif",
            svg.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StandaloneSvgPreservesExistingMarkupWhenFontReferencesNeedNoRewrite(bool referencedFont) {
        var drawing = new OfficeDrawing(100, 40);
        drawing.AddPositionedText("A", 2, 2, 90, 30, new OfficeFontInfo("Shared Font", 12));
        string withoutEmbeddedFonts = OfficeDrawingSvgExporter.ToSvg(drawing);
        drawing.Fonts.Add(referencedFont ? "Shared Font" : "Unused Font", ManagedTextShapingTestAssets.CreateFont('A'));

        string withEmbeddedFonts = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Equal(referencedFont ? 1 : 0,
            XElement.Parse(withEmbeddedFonts).Descendants().Count(element => element.Name.LocalName == "style"));
        string withoutFontDefinitions = System.Text.RegularExpressions.Regex.Replace(withEmbeddedFonts,
            "<defs><style[^>]*>.*?</style></defs>", string.Empty, System.Text.RegularExpressions.RegexOptions.Singleline);
        Assert.Equal(withoutEmbeddedFonts, withoutFontDefinitions);
    }

    [Theory]
    [InlineData("serif")]
    [InlineData("math")]
    [InlineData("Shared  Font")]
    public void StandaloneSvgKeepsLiteralQuotedEmbeddedFamilyNames(string family) {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add(family, ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30, new OfficeFontInfo(OfficeFontFamilyParser.Quote(family), 12));

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));

        Assert.Equal(OfficeFontFamilyParser.Quote(family),
            svg.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData("page-1-")]
    public void SvgKeepsLiteralAndGenericFallbacksDistinct(string prefix) {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("ABCDEF+Subset-Regular", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30,
            new OfficeFontInfo("\"serif\", ABCDEF+Subset-Regular, serif", 12));

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing, 1, OfficeSvgSizeUnit.Point, null, prefix));

        Assert.Equal("\"serif\", \"" + prefix + "ABCDEF+Subset-Regular\", serif",
            svg.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
    }

    [Theory]
    [InlineData("serif")]
    [InlineData("math")]
    public void InlineSvgScopesOnlyTheLiteralEmbeddedFamilyWhenItsNameIsGeneric(string family) {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add(family, ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30,
            new OfficeFontInfo(OfficeFontFamilyParser.Quote(family) + ", " + family, 12));

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing, 1, OfficeSvgSizeUnit.Point, null, "page-1-"));

        Assert.Equal("\"page-1-" + family + "\", " + family,
            svg.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
    }

    [Theory]
    [InlineData("", "generic(fangsong)")]
    [InlineData("page-1-", "generic(fangsong)")]
    [InlineData("page-1-", "generic(kai)")]
    [InlineData("page-1-", "generic(khmer-mul)")]
    [InlineData("page-1-", "generic(nastaliq)")]
    [InlineData("page-1-", "GeNeRiC( FANGSONG )")]
    public void SvgPreservesFunctionalGenericFallbacksWhenRewritingEmbeddedFamilies(string prefix, string fallback) {
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("ABCDEF+Subset-Regular", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30,
            new OfficeFontInfo("ABCDEF+Subset-Regular, " + fallback + ", serif", 12));

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing, 1, OfficeSvgSizeUnit.Point, null, prefix));

        Assert.Equal("\"" + prefix + "ABCDEF+Subset-Regular\", " + fallback + ", serif",
            svg.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
    }

    [Fact]
    public void InlineSvgScopesOnlyTheQuotedLiteralFamilyNamedLikeAGenericFunction() {
        const string family = "generic(fangsong)";
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add(family, ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 2, 2, 90, 30,
            new OfficeFontInfo(OfficeFontFamilyParser.Quote(family) + ", " + family, 12));

        var svg = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing, 1, OfficeSvgSizeUnit.Point, null, "page-1-"));

        Assert.Equal("\"page-1-generic(fangsong)\", generic(fangsong)",
            svg.Descendants().Single(x => x.Name.LocalName == "text").Attribute("font-family")!.Value);
    }

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
