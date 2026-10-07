using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergeLanguageContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public void DifferentChapterLanguagesRequireOptInAndRetainLocalOverrides() {
        var book = Book(); SetContext(book, "two", "ar", "rtl");
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "second-start"));
        Assert.Equal(before, book.Write().Bytes);
        book.MergeChapters("one", "two", "second-start", Options());
        var merged = EpubPublication.Load(new MemoryStream(book.Write().Bytes)).GetContentXml("one");
        Assert.Equal("en", (string?)merged.Root!.Attribute(XNamespace.Xml + "lang"));
        var wrapper = merged.Root.Element(Html + "body")!.Elements(Html + "div").Single();
        Assert.Equal("ar", (string?)wrapper.Attribute("lang"));
        Assert.Equal("ar", (string?)wrapper.Attribute(XNamespace.Xml + "lang"));
        Assert.Equal("rtl", (string?)wrapper.Attribute("dir"));
        Assert.Equal("second-start", (string?)wrapper.Elements().First().Attribute("id"));
        Assert.Equal("fr", (string?)wrapper.Descendants(Html + "span").Single(e => e.Value == "Bonjour").Attribute("lang"));
        Assert.Equal("#second", (string?)wrapper.Descendants(Html + "a").Single().Attribute("href"));
    }

    [Theory]
    [InlineData("pl", "ltr")]
    [InlineData("", "rtl")]
    public void BodyContextOverridesRootAndUnknownLanguageIsExplicit(string language, string direction) {
        var book = Book(); SetContext(book, "two", "ar", "rtl");
        var content = book.GetContentXml("two"); var body = content.Root!.Element(Html + "body")!;
        body.SetAttributeValue("lang", language); body.SetAttributeValue(XNamespace.Xml + "lang", language); body.SetAttributeValue("dir", direction);
        book.SetContentXml("two", content);
        book.MergeChapters("one", "two", "start", Options());
        var wrapper = book.GetContentXml("one").Descendants(Html + "div").Single();
        Assert.Equal(language, (string?)wrapper.Attribute("lang")); Assert.Equal(direction, (string?)wrapper.Attribute("dir"));
    }

    [Fact]
    public void AbsentDirectionDoesNotInheritTheFirstChaptersRtl() {
        var book = Book(); SetContext(book, "one", "ar", "rtl");
        book.MergeChapters("one", "two", "start", Options());
        Assert.Equal("ltr", (string?)book.GetContentXml("one").Descendants(Html + "div").Single().Attribute("dir"));
    }

    [Theory]
    [InlineData("class")]
    [InlineData("auto")]
    [InlineData("language-conflict")]
    public void UnsupportedScaffoldingStillRejectsAtomically(string kind) {
        var book = Book(); var content = book.GetContentXml("two");
        if (kind == "class") content.Root!.Element(Html + "body")!.SetAttributeValue("class", "second-only");
        if (kind == "auto") content.Root!.SetAttributeValue("dir", "auto");
        if (kind == "language-conflict") { content.Root!.SetAttributeValue("lang", "en"); content.Root.SetAttributeValue(XNamespace.Xml + "lang", "pl"); }
        book.SetContentXml("two", content); byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "start", Options()));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void WrappingKeepsSecondContainerIdentityAndRepairsItsReferences() {
        var book = EpubPublication.Create("Container merge", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<section id='part'><h1>First</h1></section>");
        book.AddChapter("two", "EPUB/two.xhtml", "Second", "<section id='part'><h1>Second</h1><p aria-labelledby='part'>Text</p><a href='#part'>Back</a></section>");
        var options = Options();
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.MergeChapters("one", "two", "start", options));
        Assert.Equal(before, book.Write().Bytes);
        options.SecondChapterIdMap = new Dictionary<string, string> { ["part"] = "second-part" };
        book.MergeChapters("one", "two", "start", options);
        var content = book.GetContentXml("one");
        Assert.Equal(new[] { "part", "second-part" }, content.Descendants(Html + "section").Select(e => (string)e.Attribute("id")!));
        Assert.Equal("second-part", (string?)content.Descendants(Html + "p").Single().Attribute("aria-labelledby"));
        Assert.Equal("#second-part", (string?)content.Descendants(Html + "a").Single().Attribute("href"));
    }

    private static EpubChapterMergeOptions Options() => new() { PreserveSecondChapterLanguageAndDirection = true };
    private static EpubPublication Book() {
        var book = EpubPublication.Create("Multilingual merge", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<h1 id='first'>First</h1>");
        book.AddChapter("two", "EPUB/two.xhtml", "Second", "<h1 id='second'>الثاني</h1><p><span lang='fr' xml:lang='fr'>Bonjour</span> <a href='#second'>Back</a></p>");
        return book;
    }
    private static void SetContext(EpubPublication book, string id, string language, string direction) {
        var content = book.GetContentXml(id); content.Root!.SetAttributeValue("lang", language);
        content.Root.SetAttributeValue(XNamespace.Xml + "lang", language); content.Root.SetAttributeValue("dir", direction);
        book.SetContentXml(id, content);
    }
}
