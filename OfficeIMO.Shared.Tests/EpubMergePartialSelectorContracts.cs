using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergePartialSelectorContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Theory]
    [InlineData("[href^='two.xhtml']")]
    [InlineData("[href$='#heading']")]
    [InlineData("[href*='xhtml']")]
    [InlineData("[href|='two.xhtml#heading']")]
    [InlineData("[href~='two.xhtml#heading']")]
    public void PartialResourceComparisonsFollowTheMatchedElement(string selector) {
        var book = Book("<a href='two.xhtml#heading'>Local</a><a href='https://example.invalid/'>External</a>");
        AddStyle(book, "a" + selector + " {color:green}");
        book.MergeChapters("one", "two", "boundary", Options());
        Assert.Equal("a[href=\"\\23 second-heading\"] {color:green}", Css(book));
    }

    [Theory]
    [InlineData("=")]
    [InlineData("~=")]
    public void OneSourceValueCanExpandIntoSeveralExactTargets(string operation) {
        var book = Book("<map id='map' name='map'></map><input name='map'/>");
        AddStyle(book, ":not([name" + operation + "map]) {color:green}");
        var options = Options();
        options.SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading", ["map"] = "second-map" };
        book.MergeChapters("one", "two", "boundary", options);
        Assert.Equal(":not(:is([name=\"map\"],[name=\"second-map\"])) {color:green}", Css(book));
    }

    [Fact]
    public void UnchangedPredicatesAndEmptySubstringOperandsRetainSourceText() {
        var book = Book("<a href='two.xhtml#heading'>Local</a><a href='https://example.invalid/'>External</a>");
        const string css = "[href^='https:'], [href$=''], [href*=''], [href^=''] {color:green}";
        AddStyle(book, css);
        book.MergeChapters("one", "two", "boundary", Options());
        Assert.Equal(css, Css(book));
    }

    [Fact]
    public void ExpansionCannotSeparateSourceMatchesThatCollapseToOneValue() {
        var book = Book("<map id='map' name='map'></map><input name='second-map'/>");
        AddStyle(book, "[name^=map] {color:green}");
        var options = Options();
        options.SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading", ["map"] = "second-map" };
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", options));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void ExpansionAboveAlternativeLimitRejectsWithoutChangingThePublication() {
        string content = string.Join("", Enumerable.Range(0, 257).Select(i => "<p id='target" + i + "'>Target</p><a href='two.xhtml#target" + i + "'>Link</a>"));
        var book = Book(content); AddStyle(book, "[href^='two.xhtml'] {color:green}");
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void Exactly256AlternativesRemainSupported() {
        string content = string.Join("", Enumerable.Range(0, 256).Select(i => "<p id='target" + i + "'>Target</p><a href='two.xhtml#target" + i + "'>Link</a>"));
        var book = Book(content); AddStyle(book, "[href^='two.xhtml'] {color:green}");
        book.MergeChapters("one", "two", "boundary", Options());
        Assert.Equal(256, Css(book).Split(new[] { "[href=" }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void ExcessiveEncodedExpansionRejectsAtomically() {
        string suffix = new string('x', 400);
        string content = string.Join("", Enumerable.Range(0, 200).Select(i => "<p id='target" + i + suffix + "'>Target</p><a href='two.xhtml#target" + i + suffix + "'>Link</a>"));
        var book = Book(content); AddStyle(book, "[href^='two.xhtml'] {color:green}");
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Book(string content) {
        var book = EpubPublication.Create("Partial selector repair", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<h1>First</h1>");
        book.AddChapter("two", "EPUB/two.xhtml", "Second", "<h1 id='heading'>Second</h1>" + content);
        return book;
    }
    private static EpubChapterMergeOptions Options() => new() {
        StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteChapterSelectors = true,
        SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading" }
    };
    private static void AddStyle(EpubPublication book, string css) {
        var content = book.GetContentXml("two"); content.Root!.Element(Html + "head")!.Add(new XElement(Html + "style", css)); book.SetContentXml("two", content);
    }
    private static string Css(EpubPublication book) => EpubPublication.Load(new MemoryStream(book.Write().Bytes))
        .GetContentXml("one").Descendants(Html + "style").Single().Value;
}
