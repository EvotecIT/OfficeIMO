using System.Threading;
using OfficeIMO.Epub;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubChapterMergeStyleContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static EpubChapterMergeOptions Append() => new() { StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles };

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void ExplicitAppendPreservesStyleOrderRebasesAssetsAndRepairsIncomingLinks(EpubVersion version) {
        var book = Book(version); byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary"));
        Assert.Equal(before, book.Write().Bytes);
        book.MergeChapters("one", "two", "boundary", Append());
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var xml = reopened.GetContentXml("one");
        var styles = xml.Root!.Element(Html + "head")!.Elements().Where(e => e.Name == Html + "style" || e.Name == Html + "link").ToArray();
        Assert.Equal(new[] { "link", "style", "link", "style" }, styles.Select(e => e.Name.LocalName));
        Assert.Equal("styles/common.css", styles[0].Attribute("href")!.Value);
        Assert.Equal("styles/common.css", styles[2].Attribute("href")!.Value);
        Assert.Equal("p{color:navy}", styles[1].Value);
        Assert.Contains("assets/art.svg", styles[3].Value);
        Assert.DoesNotContain("../assets", styles[3].Value);
        Assert.Equal("screen", styles[3].Attribute("media")!.Value);
        Assert.Equal("second-style", styles[3].Attribute("id")!.Value);
        Assert.Equal("one.xhtml#second-style", reopened.GetContentXml("other").Descendants(Html + "a").Single().Attribute("href")!.Value);
        Assert.Equal("#second", xml.Descendants(Html + "a").Single().Attribute("href")!.Value);
        Assert.Equal("caption", xml.Descendants(Html + "p").Last().Attribute("aria-describedby")!.Value);
        Assert.Equal("boundary", reopened.Read().TableOfContents[1].Fragment);
    }

    [Theory]
    [InlineData("metadata")]
    [InlineData("body")]
    [InlineData("head-attribute")]
    [InlineData("style-id")]
    [InlineData("refinement")]
    [InlineData("policy")]
    [InlineData("cancel")]
    public void AppendDoesNotDiscardOtherConflictsAndFailsAtomically(string kind) {
        var book = Book(); var second = book.GetContentXml("two");
        if (kind == "metadata") second.Root!.Element(Html + "head")!.Add(new XElement(Html + "meta", new XAttribute("name", "description"), new XAttribute("content", "Second only")));
        if (kind == "head-attribute") second.Root!.Element(Html + "head")!.SetAttributeValue("lang", "pl");
        if (kind == "body") second.Root!.Element(Html + "body")!.SetAttributeValue("class", "different");
        if (kind == "style-id") {
            var first = book.GetContentXml("one"); first.Descendants(Html + "style").Single().SetAttributeValue("id", "second-style");
            book.SetContentXml("one", first);
        }
        if (kind == "refinement") book.AddMetadataProperty("dcterms:description", "Second only", "#two");
        book.SetContentXml("two", second); byte[] before = book.Write().Bytes;
        using var cancellation = new CancellationTokenSource(); if (kind == "cancel") cancellation.Cancel();
        var options = Append(); if (kind == "policy") options.StylePolicy = (EpubChapterMergeStylePolicy)999;
        Assert.ThrowsAny<Exception>(() => book.MergeChapters("one", "two", "boundary", options, cancellation.Token));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Book(EpubVersion version = EpubVersion.Epub3) {
        var book = EpubPublication.Create("Styles", "en", version: version);
        book.AddResource("common", "EPUB/styles/common.css", "text/css", Encoding.UTF8.GetBytes("p{color:black}"));
        book.AddResource("art", "EPUB/assets/art.svg", "image/svg+xml", Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><rect width='10' height='10'/></svg>"));
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<p id='first'>First <a href='parts/two.xhtml#second'>Next</a></p>");
        book.AddChapter("two", "EPUB/parts/two.xhtml", "Two", "<h1 id='caption'>Second</h1><p id='second' aria-describedby='caption'>Second text</p>");
        book.AddChapter("other", "EPUB/other.xhtml", "Other", "<p><a href='parts/two.xhtml#second-style'>Style source</a></p>");
        foreach (string id in new[] { "one", "two" }) {
            var xml = book.GetContentXml(id); var head = xml.Root!.Element(Html + "head")!;
            head.Add(new XElement(Html + "link", new XAttribute("rel", "stylesheet"), new XAttribute("href", id == "one" ? "styles/common.css" : "../styles/common.css")));
            head.Add(id == "one" ? new XElement(Html + "style", "p{color:navy}") : new XElement(Html + "style", new XAttribute("id", "second-style"),
                new XAttribute("media", "screen"), "p{color:maroon} .art{background-image:url('../assets/art.svg')}"));
            book.SetContentXml(id, xml);
        }
        return book;
    }
}
