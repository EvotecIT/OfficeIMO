using OfficeIMO.Epub;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergeSelectorContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public void SharedStylesheetsAndImportsAreClonedWithoutRestylingOtherChapters() {
        var book = Book();
        const string shared = "@import 'nested/child.css' screen; /* keep */ #heading/**/.lead { color: navy; }";
        const string child = "@supports selector(:is(#heading)) { [id='heading'] { color: green; content: '#heading'; } }";
        book.AddStylesheet("shared", "EPUB/shared.css", shared);
        book.AddStylesheet("child", "EPUB/nested/child.css", child);
        foreach (string id in new[] { "one", "two", "three" }) {
            var content = book.GetContentXml(id);
            content.Root!.Element(Html + "head")!.Add(new XElement(Html + "link", new XAttribute("rel", "StyleSheet"), new XAttribute("href", "shared.css")));
            book.SetContentXml(id, content);
        }
        byte[] originalShared = book.GetResourceBytes("shared"), originalChild = book.GetResourceBytes("child");
        book.MergeChapters("one", "two", "boundary", Options());
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(originalShared, reopened.GetResourceBytes("shared"));
        Assert.Equal(originalChild, reopened.GetResourceBytes("child"));
        Assert.Equal("shared.css", (string?)reopened.GetContentXml("three").Descendants(Html + "link").Single().Attribute("href"));
        var merged = reopened.GetContentXml("one");
        Assert.Equal(new[] { "heading", "second.heading" }, merged.Descendants(Html + "h1").Attributes("id").Select(a => a.Value));
        Assert.Equal(new[] { "shared.css", "merge-style-1.css" }, merged.Descendants(Html + "link").Attributes("href").Select(a => a.Value));
        Assert.Equal("@import url(\"nested/merge-style-2.css\") screen; /* keep */ #second\\2e heading/**/.lead { color: navy; }",
            Encoding.UTF8.GetString(reopened.GetResourceBytes("merge-style-1")));
        string clonedChild = Encoding.UTF8.GetString(reopened.GetResourceBytes("merge-style-2"));
        Assert.Contains("selector(:is(#second\\2e heading))", clonedChild);
        Assert.Contains("[id=\"second\\2e heading\"]", clonedChild);
        Assert.Contains("content: '#heading'", clonedChild);
    }

    [Fact]
    public void EmbeddedRulesPreserveDeclarationsAndRewriteEscapedIdsAndScopedSelectors() {
        var book = Book();
        AddStyle(book, "two", "/* heading */ @scope (#heading) { @media screen { :is(#h\\65 ading, [id=heading s]) { color:#abc; content:'#heading'; } } }");
        book.MergeChapters("one", "two", "boundary", Options());
        string css = book.GetContentXml("one").Descendants(Html + "style").Single().Value;
        Assert.Contains("@scope (#second\\2e heading)", css);
        Assert.Contains(":is(#second\\2e heading, [id=\"second\\2e heading\" s])", css);
        Assert.Contains("color:#abc; content:'#heading'", css);
        book.Write();
    }

    [Theory]
    [InlineData("[id^='head'] {color:red}")]
    [InlineData("[id|=heading] {color:red}")]
    [InlineData("[id='heading' i] {color:red}")]
    [InlineData("@unknown { #heading {color:red} }")]
    public void UnsupportedSelectorsLeaveAllPackageBytesUnchanged(string css) {
        var book = Book(); AddStyle(book, "two", css);
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData("[itemref='heading'] {color:red}")]
    [InlineData("[itemref~=heading] {color:red}")]
    [InlineData("[item\\72 ef='heading'] {color:red}")]
    public void ItemReferenceSelectorsFollowReconciledTargets(string css) {
        var book = Book();
        var content = book.GetContentXml("two");
        content.Root!.Element(Html + "body")!.Add(new XElement(Html + "section",
            new XAttribute("itemscope", ""), new XAttribute("itemref", "heading"), "Referenced content"));
        book.SetContentXml("two", content);
        AddStyle(book, "two", css);
        book.MergeChapters("one", "two", "boundary", Options());
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var merged = reopened.GetContentXml("one");
        Assert.Equal("second.heading", (string?)merged.Descendants(Html + "section").Single().Attribute("itemref"));
        Assert.Contains("\"second\\2e heading\"", merged.Descendants(Html + "style").Single().Value);
    }

    [Theory]
    [InlineData("[cite~='two.xhtml#heading'] {color:red}", false)]
    [InlineData("[cite$='#heading'] {color:red}", true)]
    public void UrlAttributeSelectorsRejectBeforeChangingDocumentsOrSharedStylesheets(string css, bool linked) {
        var book = Book();
        var content = book.GetContentXml("two");
        content.Root!.Element(Html + "body")!.Add(new XElement(Html + "blockquote",
            new XAttribute("cite", "two.xhtml#heading"), "Quoted content"));
        if (linked) {
            book.AddStylesheet("shared", "EPUB/shared.css", css);
            content.Root.Element(Html + "head")!.Add(new XElement(Html + "link",
                new XAttribute("rel", "stylesheet"), new XAttribute("href", "shared.css")));
        } else content.Root.Element(Html + "head")!.Add(new XElement(Html + "style", css));
        book.SetContentXml("two", content);
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void RewritingRequiresExplicitSharedCascadeAndIsOptIn() {
        var book = Book(); AddStyle(book, "two", "#heading {color:green}");
        var options = Options(); options.StylePolicy = EpubChapterMergeStylePolicy.RequireEquivalent;
        byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.MergeChapters("one", "two", "boundary", options));
        Assert.Equal(before, book.Write().Bytes);
        options.StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles; options.RewriteSecondChapterIdSelectors = false;
        book.MergeChapters("one", "two", "boundary", options);
        Assert.Equal("#heading {color:green}", book.GetContentXml("one").Descendants(Html + "style").Single().Value);
    }

    [Fact]
    public void ClonesRespectEntryLimitsAndAvoidCaseInsensitivePathCollisions() {
        var book = Book();
        book.AddStylesheet("shared", "EPUB/shared.css", "@import 'child.css'; #heading {color:navy}");
        book.AddStylesheet("child", "EPUB/child.css", "#heading {border:1px solid}");
        book.AddStylesheet("occupied", "EPUB/MERGE-STYLE-1.CSS", "/* unrelated retained resource */");
        AddStyle(book, "two", "@import 'shared.css';");
        byte[] original = book.Write().Bytes;
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(original));
        var limited = EpubPublication.Load(new MemoryStream(original), new EpubPublicationLoadOptions { MaxEntries = archive.Entries.Count });
        byte[] before = limited.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => limited.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, limited.Write().Bytes);
        book.MergeChapters("one", "two", "boundary", Options());
        Assert.Contains("merge-style-2.css", book.GetContentXml("one").Descendants(Html + "style").Single().Value);
        Assert.Equal("/* unrelated retained resource */", Encoding.UTF8.GetString(book.GetResourceBytes("occupied")));
        book.Write();
    }

    [Fact]
    public void LanguagePrefixSelectorsRemainUnchangedAlongsideIdentifierSelectors() {
        var book = Book();
        AddStyle(book, "two", "[lang|=en] #heading {color:green} [data-kind|='chapter'] {margin:1em}");
        book.MergeChapters("one", "two", "boundary", Options());
        Assert.Equal("[lang|=en] #second\\2e heading {color:green} [data-kind|='chapter'] {margin:1em}",
            book.GetContentXml("one").Descendants(Html + "style").Single().Value);
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Selector reconciliation", "en");
        foreach (string id in new[] { "one", "two", "three" }) book.AddChapter(id, "EPUB/" + id + ".xhtml", id, "<h1 id='heading' class='lead'>" + id + "</h1>");
        return book;
    }
    private static EpubChapterMergeOptions Options() => new() {
        StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteSecondChapterIdSelectors = true,
        SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second.heading" }
    };
    private static void AddStyle(EpubPublication book, string chapter, string css) {
        var content = book.GetContentXml(chapter); content.Root!.Element(Html + "head")!.Add(new XElement(Html + "style", css)); book.SetContentXml(chapter, content);
    }
}
