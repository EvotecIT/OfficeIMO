using OfficeIMO.Epub;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergeResourceSelectorContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExactUrlSelectorsFollowMovedLinksAndCitations(bool renameIds) {
        var book = Book("<a href='two.xhtml#heading'>Local</a><blockquote cite='two.xhtml#heading'>Quote</blockquote>" +
            "<a href='https://example.invalid/book#heading'>External</a>");
        AddStyle(book, "a[href='two.xhtml#heading'], [cite='two.xhtml#heading'] {color:green} " +
            "[href] {text-decoration:underline} [href='https://example.invalid/book#heading'] {color:blue}");
        var options = Options(renameIds);
        book.MergeChapters("one", "two", "boundary", options);
        var merged = EpubPublication.Load(new MemoryStream(book.Write().Bytes)).GetContentXml("one");
        string target = renameIds ? "second-heading" : "heading";
        Assert.Equal("#" + target, (string?)merged.Descendants(Html + "a").First().Attribute("href"));
        Assert.Equal("#" + target, (string?)merged.Descendants(Html + "blockquote").Single().Attribute("cite"));
        string css = merged.Descendants(Html + "style").Single().Value;
        Assert.Contains("[href=\"\\23 " + target + "\"]", css);
        Assert.Contains("[cite=\"\\23 " + target + "\"]", css);
        Assert.Contains("[href]", css);
        Assert.Contains("[href='https://example.invalid/book#heading']", css);
    }

    [Fact]
    public void StylesheetLinkSelectorsSeeFinalClonePathsIncludingImports() {
        var book = Book("<p>Styled content</p>");
        const string child = "link[href='../shared.css'] + style {color:green}";
        book.AddStylesheet("shared", "EPUB/shared.css", "@import 'child.css'; [href='../shared.css'] {color:blue}");
        book.AddStylesheet("child", "EPUB/child.css", child);
        LinkStylesheet(book);
        AddStyle(book, "link[href='../shared.css'] {color:green}");
        book.MergeChapters("one", "two", "boundary", Options());
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var merged = reopened.GetContentXml("one");
        Assert.Equal("merge-style-1.css", (string?)merged.Descendants(Html + "link").Single().Attribute("href"));
        Assert.Contains("[href=\"merge-style-1\\2e css\"]", merged.Descendants(Html + "style").Single().Value);
        Assert.Contains("[href=\"merge-style-1\\2e css\"]", Encoding.UTF8.GetString(reopened.GetResourceBytes("merge-style-1")));
        Assert.Contains("[href=\"merge-style-1\\2e css\"]", Encoding.UTF8.GetString(reopened.GetResourceBytes("merge-style-2")));
        Assert.Equal(child, Encoding.UTF8.GetString(reopened.GetResourceBytes("child")));
    }

    [Fact]
    public void ImageAndStyleAttributeSelectorsFollowResourceRebasing() {
        var book = Book("<img src='../cover.svg' alt='Cover'/><p style=\"background-image:url('../cover.svg')\">Styled</p>");
        book.AddResource("cover", "EPUB/cover.svg", "image/svg+xml", Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><rect width='10' height='10'/></svg>"));
        AddStyle(book, "img[src='../cover.svg'] {border:1px solid} [style=\"background-image:url('../cover.svg')\"] {color:green}");
        book.MergeChapters("one", "two", "boundary", Options());
        var merged = EpubPublication.Load(new MemoryStream(book.Write().Bytes)).GetContentXml("one");
        Assert.Equal("cover.svg", (string?)merged.Descendants(Html + "img").Single().Attribute("src"));
        string css = merged.Descendants(Html + "style").Single().Value;
        Assert.Contains("[src=\"cover\\2e svg\"]", css);
        Assert.DoesNotContain("../cover.svg", css);
        Assert.Contains("[style=\"background-image", css);
    }

    [Theory]
    [InlineData("[href='../shared.css']")]
    [InlineData("[href='merge-style-1.css']")]
    public void ClonePathsCannotIntroduceAmbiguousOrAdditionalMatches(string selector) {
        var book = Book("<a href='../shared.css'>Stylesheet</a>");
        book.AddStylesheet("shared", "EPUB/shared.css", "p {color:blue}");
        LinkStylesheet(book); AddStyle(book, selector + " {color:green}");
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData("[href]")]
    [InlineData("[href='two.xhtml']")]
    public void SelectorsDependingOnRemovedBaseScaffoldingRejectAtomically(string selector) {
        var book = Book("<p>Content</p>");
        var second = book.GetContentXml("two");
        second.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", "two.xhtml")));
        book.SetContentXml("two", second); AddStyle(book, selector + " {color:green}");
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Book(string secondBody) {
        var book = EpubPublication.Create("Resource selectors", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<h1 id='first'>First</h1>");
        book.AddChapter("two", "EPUB/parts/two.xhtml", "Second", "<h1 id='heading'>Second</h1>" + secondBody);
        return book;
    }
    private static EpubChapterMergeOptions Options(bool renameIds = true) => new() {
        StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteSecondChapterIdSelectors = true,
        SecondChapterIdMap = renameIds ? new Dictionary<string, string> { ["heading"] = "second-heading" } : new Dictionary<string, string>()
    };
    private static void AddStyle(EpubPublication book, string css) {
        var content = book.GetContentXml("two"); content.Root!.Element(Html + "head")!.Add(new XElement(Html + "style", css)); book.SetContentXml("two", content);
    }
    private static void LinkStylesheet(EpubPublication book) {
        var content = book.GetContentXml("two");
        content.Root!.Element(Html + "head")!.Add(new XElement(Html + "link", new XAttribute("rel", "stylesheet"), new XAttribute("href", "../shared.css")));
        book.SetContentXml("two", content);
    }
}
