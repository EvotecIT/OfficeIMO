using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergeBodyScopeContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly XNamespace Ops = "http://www.idpf.org/2007/ops";

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void BodyIdentityAndLocalRelationshipsSurviveInSeparateScopes(EpubVersion version) {
        var book = Book(version); byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "start"));
        Assert.Equal(before, book.Write().Bytes);
        book.MergeChapters("one", "two", "start", new EpubChapterMergeOptions { PreserveBodyScopes = true });
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        XElement body = reopened.GetContentXml("one").Root!.Element(Html + "body")!;
        Assert.Null(body.Attribute("id")); Assert.Null(body.Attribute("class")); Assert.Null(body.Attribute("style"));
        XElement[] scopes = body.Elements(Html + "div").ToArray();
        Assert.Equal(new[] { "first-body", "second-body" }, scopes.Select(scope => (string?)scope.Attribute("id")));
        Assert.Equal(new[] { "first", "second" }, scopes.Select(scope => (string?)scope.Attribute("class")));
        Assert.Equal("color:navy", (string?)scopes[0].Attribute("style"));
        Assert.Equal("color:maroon", (string?)scopes[1].Attribute("style"));
        Assert.Equal("Second scope", (string?)scopes[1].Attribute("title"));
        if (version == EpubVersion.Epub3) Assert.Equal("two", (string?)scopes[1].Attribute("data-chapter"));
        Assert.Equal("start", (string?)scopes[1].Elements().First().Attribute("id"));
        if (version == EpubVersion.Epub3) Assert.Equal("second-body", (string?)scopes[1].Element(Html + "p")!.Attribute("aria-describedby"));
        Assert.Equal("#second-body", (string?)scopes[0].Descendants(Html + "a").Single().Attribute("href"));
    }

    [Fact]
    public void CollidingBodyIdsAreRenamedInsteadOfJoiningScopes() {
        var book = Book(); var second = book.GetContentXml("two");
        second.Root!.Element(Html + "body")!.SetAttributeValue("id", "first-body");
        second.Descendants(Html + "p").Single().SetAttributeValue("aria-describedby", "first-body");
        book.SetContentXml("two", second);
        var first = book.GetContentXml("one"); first.Descendants(Html + "a").Single().SetAttributeValue("href", "two.xhtml#first-body");book.SetContentXml("one", first);
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.MergeChapters("one", "two", "start", new EpubChapterMergeOptions { PreserveBodyScopes = true }));
        Assert.Equal(before, book.Write().Bytes);
        book.MergeChapters("one", "two", "start", new EpubChapterMergeOptions {
            PreserveBodyScopes = true, SecondChapterIdMap = new Dictionary<string, string> { ["first-body"] = "repaired-body" }
        });
        var body = book.GetContentXml("one").Root!.Element(Html + "body")!;
        Assert.Equal(2, body.Elements(Html + "div").Count());
        Assert.Equal("repaired-body", (string?)body.Descendants(Html + "p").Last().Attribute("aria-describedby"));
        Assert.Equal("#repaired-body", (string?)body.Descendants(Html + "a").First().Attribute("href"));
    }

    [Fact]
    public void BodyScopesComposeWithMatterAndRootLanguage() {
        var book = Book(); book.SetDocumentMatter("one", EpubDocumentMatter.FrontMatter);book.SetDocumentMatter("two", EpubDocumentMatter.BodyMatter);
        var second = book.GetContentXml("two"); second.Root!.SetAttributeValue("lang", "fr");second.Root.SetAttributeValue(XNamespace.Xml + "lang", "fr");
        second.Root.Element(Html + "body")!.SetAttributeValue("lang", "de");second.Root.Element(Html + "body")!.SetAttributeValue(XNamespace.Xml + "lang", "de");
        book.SetContentXml("two", second);
        book.MergeChapters("one", "two", "start", new EpubChapterMergeOptions {
            PreserveBodyScopes = true, PreserveDocumentMatter = true, PreserveSecondChapterLanguageAndDirection = true
        });
        XElement body = book.GetContentXml("one").Root!.Element(Html + "body")!;
        var scope = body.Descendants(Html + "div").Single(e => (string?)e.Attribute("id") == "second-body");
        Assert.Equal("de", (string?)scope.Attribute("lang"));
        Assert.Equal("bodymatter", (string?)scope.Parent!.Attribute(Ops + "type"));
        Assert.Equal("start", (string?)scope.Elements().First().Attribute("id"));
        Assert.Equal("de", (string?)scope.Parent!.Parent!.Attribute("lang"));
    }

    [Theory]
    [InlineData("role")]
    [InlineData("aria-label")]
    [InlineData("root-class")]
    public void UnresolvedSemanticsStillFailAtomically(string kind) {
        var book = Book(); var second = book.GetContentXml("two");
        (kind == "root-class" ? second.Root! : second.Root!.Element(Html + "body")!).SetAttributeValue(kind == "root-class" ? "class" : kind, kind == "role" ? "document" : "second");
        book.SetContentXml("two", second);byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "start", new EpubChapterMergeOptions { PreserveBodyScopes = true }));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void ScopedXmlIdentifiersAndInlineResourceUrlsFollowTheMovedDocument() {
        var book = Book();
        book.AddResource("cover", "EPUB/cover.svg", "image/svg+xml", System.Text.Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><rect width='10' height='10'/></svg>"));
        book.RenameResource("two", "EPUB/chapters/two.xhtml");
        var second = book.GetContentXml("two");var body = second.Root!.Element(Html + "body")!;
        body.SetAttributeValue(XNamespace.Xml + "id", "second-alias");
        body.SetAttributeValue("style", "background-image:url('../cover.svg')");
        body.Add(new XElement(Html + "a", new XAttribute("href", "#second-alias"), "Return"));
        book.SetContentXml("two", second);
        book.MergeChapters("one", "two", "start", new EpubChapterMergeOptions { PreserveBodyScopes = true });
        var scope = book.GetContentXml("one").Descendants(Html + "div").Single(e => (string?)e.Attribute("id") == "second-body");
        Assert.Equal("second-alias", (string?)scope.Attribute(XNamespace.Xml + "id"));
        Assert.Contains("cover.svg", (string?)scope.Attribute("style"));
        Assert.DoesNotContain("../", (string?)scope.Attribute("style"));
        Assert.Equal("#second-alias", (string?)scope.Element(Html + "a")!.Attribute("href"));
    }

    private static EpubPublication Book(EpubVersion version = EpubVersion.Epub3) {
        var book = EpubPublication.Create("Body scopes", "en", version: version);
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<h1>First</h1><p>Introduction.</p><a href='two.xhtml#second-body'>Second</a>");
        book.AddChapter("two", "EPUB/two.xhtml", "Second", "<h1>Second</h1><p aria-describedby='second-body'>Main text.</p>");
        foreach (var item in new[] { ("one", "first-body", "first", "color:navy"), ("two", "second-body", "second", "color:maroon") }) {
            var xml = book.GetContentXml(item.Item1);var body = xml.Root!.Element(Html + "body")!;
            body.SetAttributeValue("id", item.Item2);body.SetAttributeValue("class", item.Item3);body.SetAttributeValue("style", item.Item4);
            body.SetAttributeValue("title", item.Item1 == "one" ? "First scope" : "Second scope");
            if (version == EpubVersion.Epub3) body.SetAttributeValue("data-chapter", item.Item1);
            else body.Descendants().Attributes("aria-describedby").Remove();
            book.SetContentXml(item.Item1, xml);
        }
        return book;
    }
}
