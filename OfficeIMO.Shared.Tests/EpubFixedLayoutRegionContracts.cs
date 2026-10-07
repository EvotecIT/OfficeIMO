using OfficeIMO.Epub;
using System.Globalization;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubFixedLayoutRegionContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public void RegionsUseInvariantFractionalPixelsWithoutReorderingOrRewritingContent() {
        var book = Book(); var before = new XElement(book.GetContentXml("page").Root!.Element(Html + "body")!);
        CultureInfo original = CultureInfo.CurrentCulture;
        try {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("pl-PL");
            book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) { Regions = new[] {
                new EpubFixedLayoutRegion("second", 430.25m, 150.5m, 330m, 300m),
                new EpubFixedLayoutRegion("first", 40.125m, 150.5m, 330m, 300m)
            }});
        } finally { CultureInfo.CurrentCulture = original; }
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var document = reopened.GetContentXml("page");
        Assert.True(XNode.DeepEquals(before, document.Root!.Element(Html + "body")));
        string css = document.Descendants(Html + "style").Single().Value;
        Assert.Contains("left: 40.125px; top: 150.5px; width: 330px; height: 300px", css);
        Assert.Contains("position: absolute; box-sizing: border-box; margin: 0", css);
        Assert.Equal(new[] { "first", "second" }, document.Root.Element(Html + "body")!.Elements().Select(e => (string?)e.Attribute("id")));
        reopened.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600));
        Assert.DoesNotContain("[id=", reopened.GetContentXml("page").Descendants(Html + "style").Single().Value);
        Assert.True(XNode.DeepEquals(before, reopened.GetContentXml("page").Root!.Element(Html + "body")));
    }

    [Theory]
    [InlineData(-1, 0, 20, 20)]
    [InlineData(0, -1, 20, 20)]
    [InlineData(0, 0, 0, 20)]
    [InlineData(0, 0, 20, 0)]
    [InlineData(790, 0, 20, 20)]
    [InlineData(0, 590, 20, 20)]
    public void InvalidAndOverflowingBoxesAreRejectedAtomically(int left, int top, int width, int height) {
        var book = Book(); byte[] before = book.Write().Bytes;
        var error = Assert.Throws<InvalidDataException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) {
            Regions = new[] { new EpubFixedLayoutRegion("first", 0, 0, 10, 10), new EpubFixedLayoutRegion("second", left, top, width, height) }
        }));
        Assert.Contains("second", error.Message); Assert.Contains("800 × 600", error.Message);
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void ExtremeCoordinatesCannotOverflowArithmeticAndEdgesAreInclusive() {
        var book = Book(); byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) {
            Regions = new[] { new EpubFixedLayoutRegion("first", decimal.MaxValue, 0, decimal.MaxValue, 20) }
        }));
        Assert.Equal(before, book.Write().Bytes);
        book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) { Regions = new[] { new EpubFixedLayoutRegion("first", 790, 580, 10, 20) } });
        Assert.NotEmpty(book.Write().Bytes);
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("nested")]
    public void TargetsMustExistAtTheCanvasLevel(string target) {
        var book = Book(); byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) {
            Regions = new[] { new EpubFixedLayoutRegion(target, 0, 0, 20, 20) }
        }));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void DuplicateTargetsAndRegionCountLimitDoNotPartiallyConfigureThePage() {
        var book = Book(); byte[] before = book.Write().Bytes;
        var region = new EpubFixedLayoutRegion("first", 0, 0, 20, 20);
        Assert.Throws<InvalidDataException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) { Regions = new[] { region, region } }));
        Assert.Throws<ArgumentOutOfRangeException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) { Regions = Enumerable.Repeat(region, 1025).ToArray() }));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void SelectorEscapingRetainsPunctuationAndSupplementaryUnicodeIdentifiers() {
        const string id = "panel:\"\\😀</style>";
        var book = Book(); var doc = book.GetContentXml("page");
        doc.Root!.Element(Html + "body")!.Elements().First().SetAttributeValue("id", id);
        // This fixture has no references to the renamed outer section.
        book.SetContentXml("page", doc);
        book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) { Regions = new[] { new EpubFixedLayoutRegion(id, 0, 0, 100, 100) } });
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        string css = reopened.GetContentXml("page").Descendants(Html + "style").Single().Value;
        Assert.Contains("panel\\3a \\22 \\5c \\1f600 \\3c \\2f style\\3e ", css);
        Assert.DoesNotContain("</style>", css);
        Assert.Equal(id, (string?)reopened.GetContentXml("page").Root!.Element(Html + "body")!.Elements().First().Attribute("id"));
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Positioned regions", "en", "urn:example:regions");
        book.AddChapter("page", "EPUB/page.xhtml", "Page", "<section id='first' aria-labelledby='heading'><h1 id='heading'>First</h1><p id='nested'>Nested text</p></section><section id='second'><h2>Second</h2><a href='#heading'>Return</a></section>");
        return book;
    }
}
