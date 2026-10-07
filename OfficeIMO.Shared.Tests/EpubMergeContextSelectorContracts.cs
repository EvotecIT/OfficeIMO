using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergeContextSelectorContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Theory]
    [InlineData("map[name=map]", "map[name=\"second-map\"]")]
    [InlineData("map.lead[name^=map]", "map.lead[name=\"second-map\"]")]
    [InlineData("section > map/*keep*/[name=map]", "section > map/*keep*/[name=\"second-map\"]")]
    [InlineData("m\\61 p[name=map]", "m\\61 p[name=\"second-map\"]")]
    [InlineData(":is(map[name=map], input[name=second-map])", ":is(map[name=\"second-map\"], input[name=second-map])")]
    [InlineData("map:is(.lead,.other)[name=map]", "map:is(.lead,.other)[name=\"second-map\"]")]
    public void ExplicitElementTypesSeparateOtherwiseAmbiguousAttributeValues(string selector, string expected) {
        var book = Book(selector + " {color:green}");
        book.MergeChapters("one", "two", "boundary", Options());
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        XDocument merged = reopened.GetContentXml("one");
        Assert.Equal(expected + " {color:green}", merged.Descendants(Html + "style").Single().Value);
        Assert.Equal("second-map", (string?)merged.Descendants(Html + "map").Single().Attribute("name"));
        Assert.Equal("second-map", (string?)merged.Descendants(Html + "input").Single().Attribute("name"));
    }

    [Theory]
    [InlineData("map [name=map]")]
    [InlineData("map + [name=map]")]
    [InlineData("map, [name=map]")]
    [InlineData("map:not([name=map])")]
    [InlineData("map:is(.lead) [name=map]")]
    [InlineData("*|map[name=map]")]
    public void AmbiguousOrUnprovenContextStillRejectsAtomically(string selector) {
        var book = Book(selector + " {color:green}");
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Book(string css) {
        var book = EpubPublication.Create("Contextual selector repair", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<h1>First</h1>");
        book.AddChapter("two", "EPUB/two.xhtml", "Second", "<h1>Second</h1><section><map id='map' name='map' class='lead'></map><input name='second-map'/></section>");
        var content = book.GetContentXml("two");
        content.Root!.Element(Html + "head")!.Add(new XElement(Html + "style", css));
        book.SetContentXml("two", content);
        return book;
    }
    private static EpubChapterMergeOptions Options() => new() {
        StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteChapterSelectors = true,
        SecondChapterIdMap = new Dictionary<string, string> { ["map"] = "second-map" }
    };
}
