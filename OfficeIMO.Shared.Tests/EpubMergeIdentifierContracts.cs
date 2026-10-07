using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergeIdentifierContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public void ExplicitMapRepairsIncomingLinksLocalRelationshipsAndSvgUrls() {
        var book = Book();
        book.AddChapter("third", "EPUB/third.xhtml", "Third", "<a href='two.xhtml#heading'>Second heading</a>");
        book.AddStylesheet("external", "EPUB/external.css", "p { background: url('two.xhtml#paint'); }");
        var originalFirst = book.GetContentXml("one");
        Assert.Throws<InvalidDataException>(() => book.MergeChapters("one", "two", "second-start"));
        book.MergeChapters("one", "two", "second-start", Options());
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var merged = reopened.GetContentXml("one");
        Assert.Equal("First", merged.Descendants(Html + "h1").First().Value);
        Assert.Equal("heading", (string?)merged.Descendants(Html + "h1").First().Attribute("id"));
        Assert.Equal("second-heading", (string?)merged.Descendants(Html + "h1").Last().Attribute("id"));
        Assert.Equal("second-heading", (string?)merged.Descendants(Html + "h1").Last().Attribute(XNamespace.Xml + "id"));
        Assert.Equal("second-heading  second-description", (string?)merged.Descendants(Html + "section").Last().Attribute("aria-labelledby"));
        Assert.Equal("second-cell", (string?)merged.Descendants(Html + "td").Last().Attribute("headers"));
        Assert.Equal("#second-heading", (string?)merged.Descendants(Html + "a").Last().Attribute("href"));
        Assert.Equal("one.xhtml#second-heading", (string?)reopened.GetContentXml("third").Descendants(Html + "a").Single().Attribute("href"));
        Assert.Contains("one.xhtml#second-paint", System.Text.Encoding.UTF8.GetString(reopened.GetResourceBytes("external")));
        Assert.Contains("#second-paint", (string)merged.Descendants().Last(e => e.Name.LocalName == "rect").Attribute("fill")!);
        Assert.Equal(EpubPreflightStatus.Passed, reopened.Preflight().Checks.Single(c => c.Code == "content-identifiers").Status);
        Assert.Equal(originalFirst.Descendants(Html + "style").Single().Value, merged.Descendants(Html + "style").Single().Value);
    }

    [Theory]
    [InlineData("first-collision")]
    [InlineData("second-collision")]
    [InlineData("boundary")]
    [InlineData("missing")]
    [InlineData("shared")]
    public void InvalidMapsLeaveWholePublicationUnchanged(string failure) {
        var book = Book();
        var options = Options();
        var map = options.SecondChapterIdMap.ToDictionary(pair => pair.Key, pair => pair.Value);
        if (failure == "first-collision") map["heading"] = "description";
        if (failure == "second-collision") map["heading"] = "second-description";
        if (failure == "boundary") map["heading"] = "second-start";
        if (failure == "missing") map["absent"] = "new";
        if (failure == "shared") {
            foreach (string id in new[] { "one", "two" }) {
                var content = book.GetContentXml(id); content.Root!.Element(Html + "body")!.SetAttributeValue("id", "body"); book.SetContentXml(id, content);
            }
            map["body"] = "second-body";
        }
        options.SecondChapterIdMap = map;
        byte[] before = book.Write().Bytes;
        if (failure == "missing" || failure == "shared") Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "second-start", options));
        else Assert.Throws<InvalidDataException>(() => book.MergeChapters("one", "two", "second-start", options));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void MatchingMapNamesAndUsemapReferencesFollowRenamedIdentifiers() {
        var book = EpubPublication.Create("Maps", "en");
        string image = "<img alt='Map' src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=' usemap='#map'/><map id='map' name='map'><area alt='Target' href='#target'/></map><p id='target'>Target</p>";
        book.AddChapter("one", "EPUB/one.xhtml", "One", image);
        book.AddChapter("two", "EPUB/two.xhtml", "Two", image);
        book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions {
            SecondChapterIdMap = new Dictionary<string, string> { ["map"] = "second-map", ["target"] = "second-target" }
        });
        var content = book.GetContentXml("one");
        Assert.Equal(new[] { "map", "second-map" }, content.Descendants(Html + "map").Attributes("name").Select(a => a.Value));
        Assert.Equal(new[] { "#map", "#second-map" }, content.Descendants(Html + "img").Attributes("usemap").Select(a => a.Value));
        Assert.Equal(new[] { "#target", "#second-target" }, content.Descendants(Html + "area").Attributes("href").Select(a => a.Value));
        book.Write();
    }

    [Theory]
    [InlineData("#paint")]
    [InlineData("#%70aint")]
    [InlineData(" #paint ")]
    [InlineData("\\23 paint")]
    public void FragmentOnlyStylesheetReferencesRequireExplicitReconciliation(string target) {
        var book = Book();
        book.AddStylesheet("shared", "EPUB/shared.css", "rect { fill: url('" + target + "'); }");
        foreach (string id in new[] { "one", "two" }) {
            var content = book.GetContentXml(id);
            content.Root!.Element(Html + "head")!.Add(new XElement(Html + "link", new XAttribute("rel", "stylesheet"), new XAttribute("href", "shared.css")));
            book.SetContentXml(id, content);
        }
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "second-start", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Merge identifiers", "en");
        foreach (string id in new[] { "one", "two" }) {
            book.AddChapter(id, "EPUB/" + id + ".xhtml", id, "<h1 id='heading' xml:id='heading'>" + (id == "one" ? "First" : "Second") + "</h1><p id='description'>Description</p>" +
                "<section aria-labelledby='heading  description'><table><tr><th id='cell'>Header</th></tr><tr><td headers='cell'>Value</td></tr></table><a href='#heading'>Heading</a></section>" +
                "<svg xmlns='http://www.w3.org/2000/svg'><defs><linearGradient id='paint'/></defs><rect fill='url(#paint)' width='10' height='10'/></svg>");
            var content = book.GetContentXml(id);
            content.Root!.Element(Html + "head")!.Add(new XElement(Html + "style", "#heading, #second-heading { color: navy; }"));
            book.SetContentXml(id, content);
        }
        return book;
    }
    private static EpubChapterMergeOptions Options() => new() {
        SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading", ["description"] = "second-description", ["cell"] = "second-cell", ["paint"] = "second-paint" }
    };
}
