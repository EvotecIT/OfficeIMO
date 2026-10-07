using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergeIdSelectorContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Theory]
    [InlineData("^=", "chapter-", ":is([id=\"chapter-stable\"],[id=\"revised\"])")]
    [InlineData("$=", "old", "[id=\"revised\"]")]
    [InlineData("*=", "apter-", ":is([id=\"chapter-stable\"],[id=\"revised\"])")]
    [InlineData("|=", "chapter", ":is([id=\"chapter-stable\"],[id=\"revised\"])")]
    [InlineData("~=", "chapter-old", "[id~=\"revised\"]")]
    public void PartialIdSelectorsKeepTheirOriginalMatches(string operation, string operand, string expected) {
        var book = Book("[id" + operation + "'" + operand + "'] {color:green}");
        book.MergeChapters("one", "two", "boundary", Options());
        Assert.Equal(expected + " {color:green}", Css(book));
        var merged = EpubPublication.Load(new MemoryStream(book.Write().Bytes)).GetContentXml("one");
        Assert.Equal(new[] { "revised", "chapter-stable", "other" }, merged.Descendants(Html + "p").Attributes("id").Select(a => a.Value));
    }

    [Theory]
    [InlineData("[id^=rev]", "[id~=\"\"]")]
    [InlineData("[id=revised]", "[id~=\"\"]")]
    [InlineData("#revised", ":where([id~=\"\"])#revised")]
    public void ANewIdentifierCannotActivateAPreviouslyNonmatchingRule(string selector, string expected) {
        var book = Book(selector + " {color:red}");
        book.MergeChapters("one", "two", "boundary", Options());
        Assert.Equal(expected + " {color:red}", Css(book));
    }

    [Fact]
    public void UnchangedPredicatesAndEmptyOperandsPreserveSourceSyntax() {
        const string css = "[id$=stable], [id^=''], [id*=''], [id$=''], [id] {color:green}";
        var book = Book(css);
        book.MergeChapters("one", "two", "boundary", Options());
        Assert.Equal(css, Css(book));
    }

    [Fact]
    public void IdentifierSwapsFollowTheOriginalElements() {
        var book = Book("#chapter-old, [id=chapter-stable], [id$=old] {color:green}");
        var options = Options();
        options.SecondChapterIdMap = new Dictionary<string, string> { ["chapter-old"] = "chapter-stable", ["chapter-stable"] = "chapter-old" };
        book.MergeChapters("one", "two", "boundary", options);
        Assert.Equal("#chapter-stable, [id=\"chapter-old\"], [id=\"chapter-stable\"] {color:green}", Css(book));
    }

    private static EpubPublication Book(string css) {
        var book = EpubPublication.Create("ID selector repair", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<h1>First</h1>");
        book.AddChapter("two", "EPUB/two.xhtml", "Second", "<h1>Second</h1><p id='chapter-old'>Renamed</p><p id='chapter-stable'>Retained</p><p id='other'>Other</p>");
        var content = book.GetContentXml("two");
        content.Root!.Element(Html + "head")!.Add(new XElement(Html + "style", css));
        book.SetContentXml("two", content);
        return book;
    }
    private static EpubChapterMergeOptions Options() => new() {
        StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteChapterSelectors = true,
        SecondChapterIdMap = new Dictionary<string, string> { ["chapter-old"] = "revised" }
    };
    private static string Css(EpubPublication book) => EpubPublication.Load(new MemoryStream(book.Write().Bytes))
        .GetContentXml("one").Descendants(Html + "style").Single().Value;
}
