using System.Threading;
using OfficeIMO.Epub;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubChapterMergeContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void MergePreservesBothTocEntriesAndRepairsLinksAcrossDirectories(EpubVersion version) {
        var book = Book(version);
        book.MergeChapters("one", "two", "second-start", default);
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(new[] { "one", "other" }, reopened.Spine.Select(item => item.ManifestId));
        Assert.DoesNotContain("EPUB/parts/two.xhtml", reopened.EntryPaths);
        var content = reopened.GetContentXml("one");
        Assert.Equal("OneFirstNextSecondBackTop", content.Root!.Element(Html + "body")!.Value);
        Assert.Equal("#second", Link(content, "forward"));
        Assert.Equal("../one.xhtml#second", Link(reopened.GetContentXml("other"), "incoming"));
        Assert.Equal("#second-start", Link(content, "top"));
        Assert.Equal(new[] { "One", "Two", "Other" }, reopened.Read().TableOfContents.Select(item => item.Label));
        Assert.Equal("second-start", reopened.Read().TableOfContents[1].Fragment);
        Assert.Equal("EPUB/one.xhtml", reopened.Read().TableOfContents[1].Target);
    }

    [Fact]
    public void MergeRecombinesSplitStructuralContainersWithoutDuplicatingTheirIdentifiers() {
        var book = EpubPublication.Create("Split and merge", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<section id='shell'><div id='inner'><h1>First</h1><p id='cut'>Second <a href='#shell'>Section</a></p></div></section>");
        string original = book.GetContentXml("one").Root!.Element(Html + "body")!.Value;
        book.SplitChapter("one", "cut", "two", "EPUB/parts/two.xhtml", "Two");
        book.MergeChapters("one", "two", "second-start");
        var content = EpubPublication.Load(new MemoryStream(book.Write().Bytes)).GetContentXml("one");
        Assert.Equal(original, content.Root!.Element(Html + "body")!.Value);
        Assert.Single(content.Descendants(Html + "section"));
        Assert.Single(content.Descendants(Html + "div"));
        Assert.Equal("#second-start", content.Descendants(Html + "a").Single().Attribute("href")!.Value);
        Assert.Equal("inner", content.Descendants(Html + "span").Single().Parent!.Attribute("id")!.Value);
    }

    [Theory]
    [InlineData("style")]
    [InlineData("body")]
    [InlineData("id")]
    [InlineData("map")]
    [InlineData("order")]
    [InlineData("boundary")]
    [InlineData("cancel")]
    [InlineData("refinement")]
    [InlineData("base-target")]
    [InlineData("extra-body")]
    public void ConflictsAndCancellationLeaveThePublicationUnchanged(string kind) {
        var book = Book(); var xml = book.GetContentXml("two");
        if (kind == "style") xml.Root!.Element(Html + "head")!.Add(new XElement(Html + "style", "p{color:red}"));
        if (kind == "body") xml.Root!.Element(Html + "body")!.SetAttributeValue("class", "different");
        if (kind == "refinement") book.AddMetadataProperty("dcterms:description", "Second chapter metadata", "#two");
        if (kind == "base-target") xml.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("target", "_blank")));
        if (kind == "extra-body") xml.Root!.Add(new XElement(Html + "body", new XElement(Html + "p", "Additional content")));
        if (kind == "id") xml.Root!.Element(Html + "body")!.Add(new XElement(Html + "p", new XAttribute("id", "first"), "Duplicate"));
        if (kind == "map") {
            xml.Root!.Element(Html + "body")!.Add(new XElement(Html + "map", new XAttribute("name", "map")));
            var first = book.GetContentXml("one"); first.Root!.Element(Html + "body")!.Add(new XElement(Html + "map", new XAttribute("name", "map")));
            book.SetContentXml("one", first);
        }
        book.SetContentXml("two", xml);
        byte[] before = book.Write().Bytes;
        using var cancellation = new CancellationTokenSource(); if (kind == "cancel") cancellation.Cancel();
        Assert.ThrowsAny<Exception>(() => book.MergeChapters("one", kind == "order" ? "other" : "two", kind == "boundary" ? "second" : "second-start", cancellation.Token));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void EquivalentStylesAndHtmlBasesAreRebasedBeforeHeadComparison() {
        var book = Book(); book.AddResource("style", "EPUB/style.css", "text/css", Encoding.UTF8.GetBytes("p{color:black}"));
        foreach (string id in new[] { "one", "two" }) {
            var xml = book.GetContentXml(id);
            xml.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", id == "one" ? "one.xhtml" : "two.xhtml")),
                new XElement(Html + "link", new XAttribute("rel", "stylesheet"), new XAttribute("href", id == "one" ? "style.css" : "../style.css")));
            book.SetContentXml(id, xml);
        }
        book.MergeChapters("one", "two", "second-start");
        var content = EpubPublication.Load(new MemoryStream(book.Write().Bytes)).GetContentXml("one");
        Assert.Empty(content.Descendants(Html + "base"));
        Assert.Equal("style.css", content.Descendants(Html + "link").Single().Attribute("href")!.Value);
    }

    [Fact]
    public void MergedOriginsFollowRenameAndCannotBeMaskedByReusingTheRemovedPath() {
        var authored = EpubPublication.Create("Origins", "en");
        foreach (string id in new[] { "one", "two", "three" }) authored.AddChapter(id, "EPUB/" + id + ".xhtml", id, "<p>" + id + "</p>");
        var book = EpubPublication.Load(new MemoryStream(authored.Write().Bytes));
        book.MergeChapters("one", "two", "second-start");
        book.RenameResource("one", "EPUB/merged.xhtml");
        book.AddResource("replacement", "EPUB/two.xhtml", "text/plain", Encoding.UTF8.GetBytes("New identity"));
        var merged = book.Write().Report;
        Assert.Equal("EPUB/merged.xhtml", merged.MergedEntries["EPUB/two.xhtml"]);
        Assert.Equal("EPUB/merged.xhtml", merged.RenamedEntries["EPUB/one.xhtml"]);
        merged.RequireNoLoss();
        book.SetNavigation(new[] { new EpubNavigationEntry("Third", "EPUB/three.xhtml") });
        book.RemoveSpineItem(0); book.RemoveResource("one");
        var removed = book.Write().Report;
        Assert.Empty(removed.MergedEntries);
        Assert.Contains(removed.FidelityDiagnostics, finding => finding.Code == "EPUB_WRITE_ENTRY_REMOVED" && finding.Location == "EPUB/two.xhtml");
    }

    [Fact]
    public void CombinedChapterCannotExceedTheLoadedEntryBudget() {
        var book = EpubPublication.Create("Budget", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<p>" + new string('a', 1500) + "</p>");
        book.AddChapter("two", "EPUB/two.xhtml", "Two", "<p>" + new string('b', 1500) + "</p>");
        byte[] bytes = book.Write().Bytes;
        using var archive = new ZipArchive(new MemoryStream(bytes));
        book = EpubPublication.Load(new MemoryStream(bytes), new EpubPublicationLoadOptions { MaxEntryBytes = archive.Entries.Max(entry => entry.Length) });
        Assert.Throws<InvalidDataException>(() => book.MergeChapters("one", "two", "second-start"));
        Assert.Equal(bytes, book.Write().Bytes);
    }

    [Fact]
    public void RepeatedMergesRetainEveryOriginalResourceIdentity() {
        var authored = EpubPublication.Create("Origins", "en");
        foreach (string id in new[] { "one", "two", "three" }) authored.AddChapter(id, "EPUB/" + id + ".xhtml", id, "<p>" + id + "</p>");
        var book = EpubPublication.Load(new MemoryStream(authored.Write().Bytes));
        book.MergeChapters("two", "three", "third-start");
        book.MergeChapters("one", "two", "second-start");
        var result = book.Write(); result.Report.RequireNoLoss();
        Assert.Equal(new[] { "EPUB/three.xhtml", "EPUB/two.xhtml" }, result.Report.MergedEntries.Keys.OrderBy(value => value, StringComparer.Ordinal));
        Assert.All(result.Report.MergedEntries.Values, path => Assert.Equal("EPUB/one.xhtml", path));
        Assert.Contains("onetwothree", EpubPublication.Load(new MemoryStream(result.Bytes)).GetContentXml("one").Root!.Element(Html + "body")!.Value);
    }

    [Fact]
    public void SharedContainerAccessibilityReferenceCannotBeRedirectedToAnEmptyBoundary() {
        var book = EpubPublication.Create("Local reference", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<section id='shared'><p>First</p></section>");
        book.AddChapter("two", "EPUB/two.xhtml", "Two", "<section id='shared'><p aria-describedby='shared'>Second</p></section>");
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.MergeChapters("one", "two", "second-start"));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void RemovingAnExternalHtmlBaseCannotRetargetLinksToAnExistingLocalResource() {
        var book = EpubPublication.Create("External base", "en");
        foreach (string id in new[] { "one", "two" }) {
            book.AddChapter(id, "EPUB/" + id + ".xhtml", id, "<p><a href='appendix.xhtml'>External appendix</a></p>");
            var xml = book.GetContentXml(id);
            xml.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", "https://example.org/books/")));
            book.SetContentXml(id, xml);
        }
        book.AddChapter("appendix", "EPUB/appendix.xhtml", "Local appendix", "<p>Different local material</p>");
        book.MergeChapters("one", "two", "second-start");
        var content = EpubPublication.Load(new MemoryStream(book.Write().Bytes)).GetContentXml("one");
        Assert.All(content.Descendants(Html + "a"), link => Assert.Equal("https://example.org/books/appendix.xhtml", link.Attribute("href")!.Value));
    }

    [Theory]
    [InlineData("html")]
    [InlineData("body")]
    public void RelativeScaffoldingStylesWithDifferentTargetsRejectAtomically(string elementName) {
        var book = Book();
        book.AddResource("image-one", "EPUB/image.svg", "image/svg+xml", Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg'/>"));
        book.AddResource("image-two", "EPUB/parts/image.svg", "image/svg+xml", Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg'/>"));
        foreach (string id in new[] { "one", "two" }) {
            var xml = book.GetContentXml(id);
            var element = elementName == "html" ? xml.Root! : xml.Root!.Element(Html + "body")!;
            element.SetAttributeValue("style", "background-image:url(image.svg)");
            book.SetContentXml(id, xml);
        }
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "second-start"));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static string Link(XDocument document, string id) => document.Descendants(Html + "a").Single(element => (string?)element.Attribute("id") == id).Attribute("href")!.Value;

    private static EpubPublication Book(EpubVersion version = EpubVersion.Epub3) {
        var book = EpubPublication.Create("Merge", "en", version: version);
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<h1>One</h1><p id='first'>First<a id='forward' href='parts/two.xhtml#second'>Next</a></p>");
        book.AddChapter("two", "EPUB/parts/two.xhtml", "Two", "<p id='second'>Second<a href='../one.xhtml#first'>Back</a><a id='top' href=''>Top</a></p>");
        book.AddChapter("other", "EPUB/back/other.xhtml", "Other", "<p><a id='incoming' href='../parts/two.xhtml#second'>Incoming</a></p>");
        return book;
    }
}
