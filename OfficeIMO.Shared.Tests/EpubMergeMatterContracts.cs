using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMergeMatterContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly XNamespace Ops = "http://www.idpf.org/2007/ops";

    [Theory]
    [InlineData(EpubDocumentMatter.FrontMatter, EpubDocumentMatter.BodyMatter, "frontmatter", "bodymatter")]
    [InlineData(EpubDocumentMatter.BodyMatter, EpubDocumentMatter.BackMatter, "bodymatter", "backmatter")]
    public void MergeRetainsPartitionScopesAndNavigationBoundary(EpubDocumentMatter first, EpubDocumentMatter second,
        string firstToken, string secondToken) {
        var book = Book();
        book.SetDocumentMatter("one", first); book.SetDocumentMatter("two", second);
        byte[] original = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "start"));
        Assert.Equal(original, book.Write().Bytes);
        book.MergeChapters("one", "two", "start", new EpubChapterMergeOptions { PreserveDocumentMatter = true });
        var saved = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        XElement body = saved.GetContentXml("one").Root!.Element(Html + "body")!;
        Assert.Null(body.Attribute(Ops + "type"));
        XElement[] sections = body.Elements(Html + "section").ToArray();
        Assert.Equal(new[] { firstToken, secondToken }, sections.Select(section => (string?)section.Attribute(Ops + "type")));
        Assert.Equal("first", (string?)sections[0].Element(Html + "h1")!.Attribute("id"));
        Assert.Equal("start", (string?)sections[1].Elements().First().Attribute("id"));
        Assert.Equal("#second", (string?)sections[1].Descendants(Html + "a").Single().Attribute("href"));
    }

    [Fact]
    public void MatterScopesCombineWithEffectiveLanguageAndDirection() {
        var book = Book(); book.SetDocumentMatter("one", EpubDocumentMatter.FrontMatter);
        book.SetDocumentMatter("two", EpubDocumentMatter.BodyMatter);
        var second = book.GetContentXml("two");
        second.Root!.SetAttributeValue("lang", "ar"); second.Root.SetAttributeValue(XNamespace.Xml + "lang", "ar");
        second.Root.SetAttributeValue("dir", "rtl"); book.SetContentXml("two", second);
        book.MergeChapters("one", "two", "start", new EpubChapterMergeOptions {
            PreserveDocumentMatter = true, PreserveSecondChapterLanguageAndDirection = true
        });
        XElement wrapper = book.GetContentXml("one").Root!.Element(Html + "body")!.Element(Html + "div")!;
        Assert.Equal("ar", (string?)wrapper.Attribute("lang")); Assert.Equal("rtl", (string?)wrapper.Attribute("dir"));
        Assert.Equal("bodymatter", (string?)wrapper.Element(Html + "section")!.Attribute(Ops + "type"));
        Assert.Equal("start", (string?)wrapper.Element(Html + "section")!.Elements().First().Attribute("id"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void UnmarkedContentDoesNotInheritOtherPartition(bool firstMarked) {
        var book = Book();
        book.SetDocumentMatter(firstMarked ? "one" : "two", EpubDocumentMatter.FrontMatter);
        book.MergeChapters("one", "two", "start", new EpubChapterMergeOptions { PreserveDocumentMatter = true });
        XElement body = book.GetContentXml("one").Root!.Element(Html + "body")!;
        Assert.Null(body.Attribute(Ops + "type"));
        XElement section = Assert.Single(body.Elements(Html + "section"));
        Assert.Equal(firstMarked ? "first" : "second", (string?)section.Element(Html + "h1")!.Attribute("id"));
        Assert.Equal(firstMarked ? "second" : "first", (string?)Assert.Single(body.Elements(Html + "h1")).Attribute("id"));
    }

    [Theory]
    [InlineData("frontmatter bodymatter", "backmatter")]
    [InlineData("frontmatter introduction", "bodymatter chapter")]
    public void UnresolvedSemanticConflictsFailAtomically(string firstType, string secondType) {
        var book = Book();
        foreach (var item in new[] { ("one", firstType), ("two", secondType) }) {
            var document = book.GetContentXml(item.Item1);
            document.Root!.Element(Html + "body")!.SetAttributeValue(Ops + "type", item.Item2);
            book.SetContentXml(item.Item1, document);
        }
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "start",
            new EpubChapterMergeOptions { PreserveDocumentMatter = true }));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void ScopedContentRetainsOtherSemanticsAndRepairsCollidingReferences() {
        var book = Book();
        foreach (string id in new[] { "one", "two" }) {
            var document = book.GetContentXml(id);
            document.Root!.Element(Html + "body")!.SetAttributeValue(Ops + "type", "chapter");
            book.SetContentXml(id, document);
        }
        book.SetDocumentMatter("one", EpubDocumentMatter.FrontMatter);
        book.SetDocumentMatter("two", EpubDocumentMatter.BodyMatter);
        var second = book.GetContentXml("two");
        second.Descendants(Html + "h1").Single().SetAttributeValue("id", "first");
        second.Descendants(Html + "a").Single().SetAttributeValue("href", "#first");
        book.SetContentXml("two", second);
        book.MergeChapters("one", "two", "start", new EpubChapterMergeOptions {
            PreserveDocumentMatter = true, SecondChapterIdMap = new Dictionary<string, string> { ["first"] = "repaired" }
        });
        XElement body = book.GetContentXml("one").Root!.Element(Html + "body")!;
        Assert.Equal("chapter", (string?)body.Attribute(Ops + "type"));
        XElement section = body.Elements(Html + "section").Last();
        Assert.Equal("repaired", (string?)section.Element(Html + "h1")!.Attribute("id"));
        Assert.Equal("#repaired", (string?)section.Descendants(Html + "a").Single().Attribute("href"));
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Matter merge", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "First", "<h1 id='first'>First</h1><p>Introduction.</p>");
        book.AddChapter("two", "EPUB/two.xhtml", "Second", "<h1 id='second'>Second</h1><p>Main text.</p><a href='#second'>Back</a>");
        return book;
    }
}
