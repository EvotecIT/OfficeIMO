using OfficeIMO.Epub;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergeNestedSelectorContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NestedSelectorsAndInterleavedDeclarationsKeepSourceOrder(bool external) {
        const string css = """
            .chapter {
              color: #abc;
              & > #heading, h1#heading { padding:1em; }
              h1:is(#heading) { font-weight:bold; }
              @media screen {
                margin:1em;
                [id^='head'] { color:green; @supports (display:block) { border:1px solid; } }
                padding:2em;
              }
              background:white;
            }
            """;
        var book = Book(); AddStyles(book, css, external);
        book.MergeChapters("one", "two", "boundary", Options());
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        string result = external ? Encoding.UTF8.GetString(reopened.GetResourceBytes("merge-style-1")) :
            reopened.GetContentXml("one").Descendants(Html + "style").Single().Value;
        Assert.Equal(css.Replace("#heading", "#second-heading").Replace("[id^='head']", "[id=\"second-heading\"]"), result);
        if (external) Assert.Equal(css, Encoding.UTF8.GetString(reopened.GetResourceBytes("styles")));
    }

    [Fact]
    public void CustomPropertyBlocksAndFunctionsAreDataRatherThanNestedSelectors() {
        const string css = """
            .chapter {
              --payload: { #heading { content:'#heading'; value:[#heading]; } } tail;
              --function: fn({nested:[#heading]});
              --quoted: "#heading { }";
              & #heading { color:#abc; content:'#heading'; --nested:{#heading}; }
              --after: {#heading};
            }
            """;
        var book = Book(); AddStyles(book, css, false);
        book.MergeChapters("one", "two", "boundary", Options());
        Assert.Equal(css.Replace("& #heading", "& #second-heading"), book.GetContentXml("one").Descendants(Html + "style").Single().Value);
    }

    [Theory]
    [InlineData(".chapter { @unknown { #heading {color:red} } }")]
    public void UnsupportedNestedSyntaxStillFailsAtomically(string css) {
        var book = Book(); AddStyles(book, css, false); byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void ExcessiveNestingLeavesTheEditableDraftUnchanged() {
        var book = Book();
        AddStyles(book, string.Concat(Enumerable.Repeat("& {", 65)) + "#heading {color:red}" + new string('}', 65), false);
        string package = book.GetPackageXml().ToString();
        byte[] before = book.GetResourceBytes("two");
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(package, book.GetPackageXml().ToString());
        Assert.Equal(before, book.GetResourceBytes("two"));
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Nested selectors", "en");
        foreach (string id in new[] { "one", "two" }) book.AddChapter(id, "EPUB/" + id + ".xhtml", id,
            "<section class='chapter'><h1 id='heading'>" + id + "</h1></section>");
        return book;
    }
    private static EpubChapterMergeOptions Options() => new() {
        StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteChapterSelectors = true,
        SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading" }
    };
    private static void AddStyles(EpubPublication book, string css, bool external) {
        var content = book.GetContentXml("two");
        if (external) {
            book.AddStylesheet("styles", "EPUB/styles.css", css);
            content.Root!.Element(Html + "head")!.Add(new XElement(Html + "link", new XAttribute("rel", "stylesheet"), new XAttribute("href", "styles.css")));
        } else content.Root!.Element(Html + "head")!.Add(new XElement(Html + "style", css));
        book.SetContentXml("two", content);
    }
}
