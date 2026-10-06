using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergeRelationshipSelectorContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public void ExactAndTokenSelectorsFollowActualReferenceChangesAndRetainWhitespace() {
        var book = Book("<section aria-labelledby='heading&#9;other'><p id='other'>Other</p></section>");
        AddStyle(book, "section[aria-labelledby='heading\\9 other'] {color:green} " +
            "[aria-labelledby~=heading] {border:1px solid} [aria-labelledby] {padding:1em} " +
            "[aria-labelledby~='heading other'] {color:red} [aria-labelledby~=''] {color:red}");
        book.MergeChapters("one", "two", "boundary", Options());
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var merged = reopened.GetContentXml("one");
        Assert.Equal("second-heading\tother", (string?)merged.Descendants(Html + "section").Single().Attribute("aria-labelledby"));
        string css = merged.Descendants(Html + "style").Single().Value;
        Assert.Contains("[aria-labelledby=\"second-heading\\9 other\"]", css);
        Assert.Contains("[aria-labelledby~=\"second-heading\"]", css);
        Assert.Contains("[aria-labelledby]", css);
        Assert.Contains("[aria-labelledby~='heading other']", css);
        Assert.Contains("[aria-labelledby~='']", css);
    }

    [Fact]
    public void LabelAndNamedMapSelectorsUseRewrittenValues() {
        var book = Book("<input id='field'/><label for='field'>Field</label>" +
            "<map id='map' name='map'></map><object usemap='#map'>Map fallback</object>");
        AddStyle(book, "label[for=field s], map[name=map], object[usemap='#map'] {color:green}");
        var options = Options();
        options.SecondChapterIdMap = new Dictionary<string, string> {
            ["heading"] = "second-heading", ["field"] = "second-field", ["map"] = "second-map"
        };
        book.MergeChapters("one", "two", "boundary", options);
        var merged = EpubPublication.Load(new MemoryStream(book.Write().Bytes)).GetContentXml("one");
        string css = merged.Descendants(Html + "style").Single().Value;
        Assert.Contains("[for=\"second-field\" s]", css);
        Assert.Contains("[name=\"second-map\"]", css);
        Assert.Contains("[usemap=\"\\23 second-map\"]", css);
        Assert.Equal("second-field", (string?)merged.Descendants(Html + "label").Single().Attribute("for"));
        Assert.Equal("#second-map", (string?)merged.Descendants(Html + "object").Single().Attribute("usemap"));
    }

    [Theory]
    [InlineData("map", "[name=map]")]
    [InlineData("second-map", "[name=map]")]
    [InlineData("second-map", "[name=second-map]")]
    [InlineData("map", "[name~=map]")]
    public void ConflictingOrdinaryAttributeValuesRejectAtomically(string inputName, string selector) {
        var book = Book("<map id='map' name='map'></map><input name='" + inputName + "'/>");
        AddStyle(book, selector + " {color:green}");
        var options = Options();
        options.SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading", ["map"] = "second-map" };
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", options));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData("[aria-labelledby^=head]")]
    [InlineData("[aria-labelledby=heading i]")]
    [InlineData("[ARIA-LABELLEDBY=heading]")]
    public void UnsupportedRelationshipOperatorsAndNameCasingRejectAtomically(string selector) {
        var book = Book("<section aria-labelledby='heading'>Labelled content</section>");
        AddStyle(book, selector + " {color:green}");
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "boundary", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void SelectorsWithoutMatchingReferencesRemainUnchanged() {
        var book = Book("<section aria-labelledby='heading'>Labelled content</section>");
        const string css = "[aria-labelledby=absent], [headers~=absent], [itemref] {color:green}";
        AddStyle(book, css);
        book.MergeChapters("one", "two", "boundary", Options());
        Assert.Equal(css, book.GetContentXml("one").Descendants(Html + "style").Single().Value);
    }

    private static EpubPublication Book(string secondBody) {
        var book = EpubPublication.Create("Relationship selectors", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<h1 id='heading'>First</h1>");
        book.AddChapter("two", "EPUB/two.xhtml", "Second", "<h1 id='heading'>Second</h1>" + secondBody);
        return book;
    }
    private static EpubChapterMergeOptions Options() => new() {
        StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles, RewriteSecondChapterIdSelectors = true,
        SecondChapterIdMap = new Dictionary<string, string> { ["heading"] = "second-heading" }
    };
    private static void AddStyle(EpubPublication book, string css) {
        var content = book.GetContentXml("two");
        content.Root!.Element(Html + "head")!.Add(new XElement(Html + "style", css));
        book.SetContentXml("two", content);
    }
}
