using System.Threading;
using OfficeIMO.Epub;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubStructureAuthoringContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly XNamespace Ops = "http://www.idpf.org/2007/ops";

    [Fact]
    public void MatterChangesOnlyItsPartitionAndPreservesReadingOrder() {
        var book = Book();
        XDocument xml = book.GetContentXml("chapter");
        xml.Root!.Element(Html + "body")!.SetAttributeValue(Ops + "type", "frontmatter preface");
        book.SetContentXml("chapter", xml);
        book.SetDocumentMatter("chapter", EpubDocumentMatter.BodyMatter);
        Assert.Equal("preface bodymatter", (string?)book.GetContentXml("chapter").Root!.Element(Html + "body")!.Attribute(Ops + "type"));
        book.SetDocumentMatter("chapter", EpubDocumentMatter.BackMatter);
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal("preface backmatter", (string?)reopened.GetContentXml("chapter").Root!.Element(Html + "body")!.Attribute(Ops + "type"));
        Assert.Equal(new[] { "chapter" }, reopened.Spine.Select(item => item.ManifestId));
        byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentOutOfRangeException>(() => book.SetDocumentMatter("chapter", (EpubDocumentMatter)99));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("assets/")]
    [InlineData("../")]
    public void PageMarkersPreserveNavigationAndResolveUnderHtmlBases(string? htmlBase) {
        var book = Book();
        if (htmlBase != null) {
            XDocument nav = book.GetContentXml("navigation");
            nav.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", htmlBase)));
            // Existing TOC links must be valid under the edited base as well.
            nav.Descendants(Html + "a").Single().SetAttributeValue("href", htmlBase == "../" ? "EPUB/text/chapter.xhtml" : "../text/chapter.xhtml");
            book.SetContentXml("navigation", nav);
        }
        book.AddPrintPageMarker("chapter", "p-one", "iv", "Print pages");
        XDocument navigation = book.GetContentXml("navigation");
        XElement pageList = navigation.Descendants(Html + "nav").Single(e => (string?)e.Attribute(Ops + "type") == "page-list");
        pageList.SetAttributeValue("data-retained", "yes");
        book.SetContentXml("navigation", navigation);
        book.AddPrintPageMarker("chapter", "p-two", "1", "Ignored for existing heading");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(new[] { "iv", "1" }, reopened.Read().PageList.Select(page => page.Label));
        Assert.All(reopened.Read().PageList, page => Assert.Equal("EPUB/text/chapter.xhtml", page.Target));
        XElement marker = reopened.GetContentXml("chapter").Descendants(Html + "span").First();
        Assert.Equal("doc-pagebreak", (string?)marker.Attribute("role"));
        Assert.Equal("iv", (string?)marker.Attribute("aria-label"));
        Assert.Equal("pagebreak", (string?)marker.Attribute(Ops + "type"));
        XElement retained = reopened.GetContentXml("navigation").Descendants(Html + "nav").Single(e => (string?)e.Attribute(Ops + "type") == "page-list");
        Assert.Equal("Print pages", retained.Element(Html + "h1")!.Value);
        Assert.Equal("yes", (string?)retained.Attribute("data-retained"));
        Assert.Single(reopened.Read().TableOfContents);
        byte[] before = reopened.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => reopened.AddPrintPageMarker("chapter", "p-one", "duplicate"));
        Assert.Equal(before, reopened.Write().Bytes);
    }

    [Fact]
    public void FailedNavigationBudgetAndCancellationDoNotMutateTheMarker() {
        byte[] input = Book().Write().Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxExpandedBytes = zip.Entries.Sum(entry => entry.Length) + 100 });
        Assert.Throws<InvalidDataException>(() => limited.AddPrintPageMarker("chapter", "p-one", "1"));
        Assert.Equal(input, limited.Write().Bytes);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => limited.AddPrintPageMarker("chapter", "p-one", "1", cancellationToken: cancellation.Token));
        Assert.Equal(input, limited.Write().Bytes);
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Print pages", "en");
        book.AddChapter("chapter", "EPUB/text/chapter.xhtml", "Chapter", "<h1>Chapter</h1><p><span id='p-one'/>First page.</p><p><span id='p-two'/>Second page.</p>");
        return book;
    }
}
