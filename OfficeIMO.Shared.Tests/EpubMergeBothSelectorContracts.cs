using OfficeIMO.Epub;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergeBothSelectorContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public void FirstChapterSelectorsFollowIncomingLinksWithoutRenamingItsIds() {
        var book = Book();
        AddStyle(book, "one", "#heading {color:blue} [href='two.xhtml#heading'] {border:1px solid}");
        AddStyle(book, "two", "#heading {color:green} [href='#heading'] {border:2px solid}");
        book.MergeChapters("one", "two", "boundary", Options());
        var merged = EpubPublication.Load(new MemoryStream(book.Write().Bytes)).GetContentXml("one");
        string[] styles = merged.Descendants(Html + "style").Select(e => e.Value).ToArray();
        Assert.Equal("#heading {color:blue} [href=\"\\23 second-heading\"] {border:1px solid}", styles[0]);
        Assert.Equal("#second-heading {color:green} [href=\"\\23 second-heading\"] {border:2px solid}", styles[1]);
        Assert.Equal(new[] { "heading", "second-heading" }, merged.Descendants(Html + "h1").Attributes("id").Select(a => a.Value));
        Assert.All(merged.Descendants(Html + "a"), a => Assert.Equal("#second-heading", (string?)a.Attribute("href")));
    }

    [Fact]
    public void SharedCssHasSeparateChapterCopiesAndOneCombinedAllocationBudget() {
        var book = Book();
        const string original = "#heading {color:blue} [href='two.xhtml#heading'] {border:1px solid}";
        book.AddStylesheet("shared", "EPUB/shared.css", original);
        foreach (string id in new[] { "one", "two", "three" }) Link(book, id);
        byte[] bytes = book.Write().Bytes;
        using var zip = new ZipArchive(new MemoryStream(bytes));
        var limited = EpubPublication.Load(new MemoryStream(bytes), new EpubPublicationLoadOptions { MaxEntries = zip.Entries.Count });
        byte[] before = limited.Write().Bytes;
        // One private copy fits after removing the second chapter; two copies do not.
        Assert.Throws<InvalidDataException>(() => limited.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, limited.Write().Bytes);

        book.MergeChapters("one", "two", "boundary", Options());
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        string[] links = reopened.GetContentXml("one").Descendants(Html + "link").Attributes("href").Select(a => a.Value).ToArray();
        Assert.Equal(2, links.Length); Assert.NotEqual(links[0], links[1]);
        string Css(string href) => Encoding.UTF8.GetString(reopened.GetResourceBytes(
            reopened.Manifest.Single(item => item.Reference.ContainerPath == "EPUB/" + href).Id));
        Assert.Contains("#heading {color:blue}", Css(links[0]));
        Assert.Contains("[href=\"\\23 second-heading\"]", Css(links[0]));
        Assert.Contains("#second-heading {color:blue}", Css(links[1]));
        Assert.Contains("[href='two.xhtml#heading']", Css(links[1]));
        Assert.Equal("shared.css", (string?)reopened.GetContentXml("three").Descendants(Html + "link").Single().Attribute("href"));
        Assert.Equal(original, Encoding.UTF8.GetString(reopened.GetResourceBytes("shared")));
    }

    [Theory]
    [InlineData("[href$='#heading'] {color:red}")]
    [InlineData("[href='#second-heading'] {color:red}")]
    public void FirstChapterConflictsRejectTheWholeMerge(string css) {
        var book = Book(); AddStyle(book, "one", css);
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Both chapter selectors", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<h1 id='heading'>First</h1><a href='two.xhtml#heading'>Continue</a>");
        book.AddChapter("two", "EPUB/two.xhtml", "Second", "<h1 id='heading'>Second</h1><a href='#heading'>Local</a>");
        book.AddChapter("three", "EPUB/three.xhtml", "Third", "<h1>Third</h1>");
        return book;
    }
    private static EpubChapterMergeOptions Options() => new() {
        StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteChapterSelectors = true,
        SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading" }
    };
    private static void AddStyle(EpubPublication book, string id, string css) {
        var content = book.GetContentXml(id); content.Root!.Element(Html + "head")!.Add(new XElement(Html + "style", css)); book.SetContentXml(id, content);
    }
    private static void Link(EpubPublication book, string id) {
        var content = book.GetContentXml(id); content.Root!.Element(Html + "head")!.Add(new XElement(Html + "link",
            new XAttribute("rel", "stylesheet"), new XAttribute("href", "shared.css"))); book.SetContentXml(id, content);
    }
}
