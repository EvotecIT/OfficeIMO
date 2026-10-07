using System.Threading;
using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubChapterSplitContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void NestedSplitRepairsIncomingAndOutgoingLinksAndPreservesReadingOrder(EpubVersion version) {
        var book = Book(version);
        book.SplitChapter("one", "cut", "second", "EPUB/parts/second.xhtml", "Second half");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(new[] { "one", "second", "other" }, reopened.Spine.Select(item => item.ManifestId));
        var first = reopened.GetContentXml("one"); var second = reopened.GetContentXml("second");
        Assert.DoesNotContain(first.Descendants(), node => (string?)node.Attribute("id") == "cut");
        Assert.DoesNotContain(second.Descendants(), node => (string?)node.Attribute("id") == "before");
        Assert.Equal("Second half", second.Descendants(Html + "title").Single().Value);
        Assert.Equal("chapter", second.Descendants(Html + "section").Single().Attribute("class")!.Value);
        Assert.Equal("parts/second.xhtml#cut", Link(first, "forward"));
        Assert.Equal("../one.xhtml#before", Link(second, "back"));
        Assert.Equal("#shell", Link(second, "local"));
        Assert.Equal("parts/second.xhtml?mode=1#cut", Link(reopened.GetContentXml("other"), "incoming"));
        Assert.Equal("one.xhtml", Link(reopened.GetContentXml("other"), "whole"));
        Assert.Equal(new[] { "EPUB/one.xhtml", "EPUB/parts/second.xhtml", "EPUB/other.xhtml" }, reopened.Read().TableOfContents.Select(item => item.Target));
        Assert.Equal("cut", reopened.Read().TableOfContents[1].Fragment);
    }

    [Fact]
    public void CopiedHeadRebasesStylesheetAndFragmentCssWithoutChangingOriginalDocumentLinks() {
        var book = Book();
        book.AddResource("css", "EPUB/style.css", "text/css", System.Text.Encoding.UTF8.GetBytes("p{color:black}"));
        var xml = book.GetContentXml("one");
        xml.Root!.Element(Html + "head")!.Add(new XElement(Html + "link", new XAttribute("rel", "stylesheet"), new XAttribute("href", "style.css")),
            new XElement(Html + "style", "p { filter:url(#cut); }"));
        book.SetContentXml("one", xml);
        book.SplitChapter("one", "cut", "second", "EPUB/parts/second.xhtml", "Second");
        Assert.Equal("../style.css", book.GetContentXml("second").Descendants(Html + "link").Single().Attribute("href")!.Value);
        Assert.Contains("parts/second.xhtml#cut", book.GetContentXml("one").Descendants(Html + "style").Single().Value);
        Assert.Contains("url(#cut)", book.GetContentXml("second").Descendants(Html + "style").Single().Value);
        book.Write();
    }

    [Theory]
    [InlineData("idref")]
    [InlineData("inline")]
    [InlineData("empty")]
    [InlineData("collision")]
    [InlineData("ancestor-path")]
    [InlineData("cancel")]
    public void UnsupportedCutsAreAtomic(string kind) {
        var book = Book();
        var xml = book.GetContentXml("one");
        if (kind == "idref") xml.Descendants(Html + "p").First().SetAttributeValue("aria-describedby", "cut");
        book.SetContentXml("one", xml);
        byte[] before = book.Write().Bytes;
        using var cancellation = new CancellationTokenSource(); if (kind == "cancel") cancellation.Cancel();
        Assert.ThrowsAny<Exception>(() => book.SplitChapter("one", kind == "inline" ? "forward" : kind == "empty" ? "shell" : "cut", "second",
            kind == "collision" ? "EPUB/OTHER.xhtml" : kind == "ancestor-path" ? "EPUB/one.xhtml/second.xhtml" : "EPUB/parts/second.xhtml", "Second", cancellation.Token));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void HtmlBaseRetainsAssetResolutionWhileMovedFragmentsFollowTheNewChapter() {
        var book = Book(); var xml = book.GetContentXml("one");
        xml.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", "one.xhtml")));
        book.SetContentXml("one", xml);
        book.SplitChapter("one", "cut", "second", "EPUB/parts/second.xhtml", "Second");
        var second = book.GetContentXml("second");
        string baseHref = second.Descendants(Html + "base").Single().Attribute("href")!.Value;
        Assert.Equal("../one.xhtml", baseHref);
        Assert.Equal("EPUB/parts/second.xhtml", EpubReference.Resolve("EPUB/parts/second.xhtml", baseHref, Link(second, "local")).ContainerPath);
        Assert.Equal("EPUB/one.xhtml", EpubReference.Resolve("EPUB/parts/second.xhtml", baseHref, Link(second, "back")).ContainerPath);
        book.Write();
    }

    [Fact]
    public void RetentionFailureLeavesTheLoadedPublicationUnchanged() {
        byte[] bytes = Book().Write().Bytes;
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        var book = EpubPublication.Load(new MemoryStream(bytes), new EpubPublicationLoadOptions { MaxEntryBytes = archive.Entries.Max(entry => entry.Length) });
        Assert.Throws<InvalidDataException>(() => book.SplitChapter("one", "cut", "second", "EPUB/parts/second.xhtml", new string('x', 5000)));
        Assert.Equal(bytes, book.Write().Bytes);
    }

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void NestedNavigationAndPageTargetsAreRepairedWithoutFlattening(EpubVersion version) {
        var book = Book(version);
        book.SetNavigation(new[] { new EpubNavigationEntry("Chapter", "EPUB/one.xhtml#cut", new[] {
            new EpubNavigationEntry("Earlier", "EPUB/one.xhtml#before"), new EpubNavigationEntry("Later", "EPUB/one.xhtml#cut") }) },
            new[] { new EpubNavigationEntry("2", "EPUB/one.xhtml#cut") });
        book.SplitChapter("one", "cut", "second", "EPUB/parts/second.xhtml", "Second");
        var read = EpubPublication.Load(new MemoryStream(book.Write().Bytes)).Read();
        Assert.Equal("EPUB/one.xhtml", read.TableOfContents[0].Target);
        Assert.Null(read.TableOfContents[0].Fragment);
        Assert.Equal(2, read.TableOfContents[0].Children.Count);
        Assert.Equal("EPUB/parts/second.xhtml", read.TableOfContents[0].Children[1].Target);
        Assert.Equal("cut", read.PageList.Single().Fragment);
        Assert.Equal("EPUB/parts/second.xhtml", read.PageList.Single().Target);
    }

    [Fact]
    public void FixedLayoutCannotBeSplitAsReflowableContent() {
        var book = Book(); book.AddMetadataProperty("rendition:layout", "pre-paginated");
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.SplitChapter("one", "cut", "second", "EPUB/parts/second.xhtml", "Second"));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData("<p><label for='field'>Name</label></p><div id='cut'><input id='field'/></div>")]
    [InlineData("<p><input list='choices'/></p><div id='cut'><datalist id='choices'><option value='One'/></datalist></div>")]
    [InlineData("<p><img src='data:image/png;base64,iVBORw0KGgo=' alt='Map' usemap='#map'/></p><div id='cut'><map id='map' name='map'><area href='#cut' alt='Section'/></map></div>")]
    public void DocumentLocalHtmlRelationshipsCannotBeSeparated(string content) {
        var book = EpubPublication.Create("Local references", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "One", content);
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.SplitChapter("one", "cut", "second", "EPUB/second.xhtml", "Second"));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageMapNameRemainsLocalUnderARebasedHtmlBase(bool rename) {
        var book = EpubPublication.Create("Map", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<h1>First</h1><div id='cut'><img src='data:image/png;base64,iVBORw0KGgo=' alt='Map' usemap='#map'/><map id='map' name='map'><area href='#cut' alt='Section'/></map></div>");
        var xml = book.GetContentXml("one");
        xml.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", "one.xhtml")));
        book.SetContentXml("one", xml);
        if (rename) book.RenameResource("one", "EPUB/parts/one.xhtml");
        else book.SplitChapter("one", "cut", "second", "EPUB/parts/second.xhtml", "Second");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal("#map", reopened.GetContentXml(rename ? "one" : "second").Descendants(Html + "img").Single().Attribute("usemap")!.Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("#")]
    [InlineData("one.xhtml")]
    public void WholeDocumentLinksMovedIntoSecondChapterRetainOriginalDestination(string href) {
        var book = EpubPublication.Create("Links", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<p>First</p><div id='cut'><a href='" + href + "'>Chapter start</a></div>");
        book.SplitChapter("one", "cut", "second", "EPUB/parts/second.xhtml", "Second");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        string rewritten = reopened.GetContentXml("second").Descendants(Html + "a").Single().Attribute("href")!.Value;
        Assert.Equal("EPUB/one.xhtml", EpubReference.Resolve("EPUB/parts/second.xhtml", rewritten).ContainerPath);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void EmptySvgAndImageMapLinksRetainDestinationWithOrWithoutHtmlBase(bool svg, bool withBase) {
        var book = EpubPublication.Create("Links", "en");
        string link = svg ? "<svg xmlns='http://www.w3.org/2000/svg'><a href=''><text>Start</text></a></svg>" :
            "<map name='map' id='map'><area href='' alt='Start'/></map>";
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<p>First</p><div id='cut' style=\"background-image:url('')\">" + link + "</div>");
        if (withBase) {
            var original = book.GetContentXml("one");
            original.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", "one.xhtml")));
            book.SetContentXml("one", original);
        }
        book.SplitChapter("one", "cut", "second", "EPUB/parts/second.xhtml", "Second");
        var content = EpubPublication.Load(new MemoryStream(book.Write().Bytes)).GetContentXml("second");
        string href = content.Descendants().Single(element => element.Name.LocalName == (svg ? "a" : "area")).Attribute("href")!.Value;
        var owner = new Uri("https://epub.test/EPUB/parts/second.xhtml");
        string? baseHref = content.Descendants(Html + "base").SingleOrDefault()?.Attribute("href")?.Value;
        Assert.Equal("/EPUB/one.xhtml", new Uri(baseHref == null ? owner : new Uri(owner, baseHref), href).AbsolutePath);
        Assert.Equal("background-image:url('')", content.Descendants(Html + "div").Single().Attribute("style")!.Value);
    }

    private static string Link(XDocument document, string id) => document.Descendants(Html + "a").Single(node => (string?)node.Attribute("id") == id).Attribute("href")!.Value;

    private static EpubPublication Book(EpubVersion version = EpubVersion.Epub3) {
        var book = EpubPublication.Create("Split", "en", version: version);
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<section id='shell' class='chapter'><p id='before'>Before <a id='forward' href='#cut'>next</a></p><h2 id='cut'>Second</h2><p>After <a id='back' href='#before'>back</a><a id='local' href='#shell'>local</a></p></section>");
        book.AddChapter("other", "EPUB/other.xhtml", "Other", "<p><a id='incoming' href='one.xhtml?mode=1#cut'>Incoming</a><a id='whole' href='one.xhtml'>Whole</a></p>");
        return book;
    }
}
