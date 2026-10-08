using System.Threading;
using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubScopedEditContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    [Fact]
    public void BatchCanReplaceTargetsAndIncomingLinksTogetherWithoutChangingSurroundingContent() {
        var book = Book();
        var target = Element(book, "one", "target"); var link = Element(book, "two", "link");
        var replacement = new XElement(target); replacement.SetAttributeValue("id", "updated"); replacement.Value = "Edited";
        var newLink = new XElement(link); newLink.SetAttributeValue("href", "one.xhtml#updated");
        var edits = new[] { new EpubContentEdit("one", "target", target, replacement), new EpubContentEdit("two", "link", link, newLink) };
        replacement.Value = "Later mutation"; target.Value = "Later mutation";
        edits[0].Replacement!.Value = "Getter mutation";
        book.ApplyContentEdits(edits);
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal("Edited", Element(reopened, "one", "updated").Value);
        Assert.Equal("Keep", Element(reopened, "one", "keep").Value);
        Assert.Equal("one.xhtml#updated", Element(reopened, "two", "link").Attribute("href")!.Value);
    }
    [Theory]
    [InlineData("stale")]
    [InlineData("overlap")]
    [InlineData("duplicate")]
    [InlineData("broken-link")]
    [InlineData("scaffolding")]
    [InlineData("cancel")]
    public void RejectedEditsLeaveTheWholePublicationUnchanged(string kind) {
        var book = Book(); var target = Element(book, "one", "target"); var replacement = new XElement(target); replacement.Value = "Edited";
        var edits = new List<EpubContentEdit>();
        if (kind == "stale") target.Value = "Stale";
        edits.Add(new EpubContentEdit("one", "target", target, kind == "broken-link" ? null : replacement));
        if (kind == "duplicate") edits.Add(edits[0]);
        if (kind == "overlap") { var parent = Element(book, "one", "section"); edits.Add(new EpubContentEdit("one", "section", parent, null)); }
        if (kind == "scaffolding") {
            var xml = book.GetContentXml("one"); xml.Root!.SetAttributeValue("id", "root"); book.SetContentXml("one", xml);
            edits.Add(new EpubContentEdit("one", "root", xml.Root, xml.Root));
        }
        byte[] before = book.Write().Bytes;
        using var token = new CancellationTokenSource(); if (kind == "cancel") token.Cancel();
        Assert.ThrowsAny<Exception>(() => book.ApplyContentEdits(edits, token.Token));
        Assert.Equal(before, book.Write().Bytes);
    }
    [Fact]
    public void DeletionRetainsUnselectedSiblings() {
        var book = Book(); var keep = Element(book, "one", "keep");
        book.ApplyContentEdits(new[] { new EpubContentEdit("one", "keep", keep, null) });
        Assert.DoesNotContain(book.GetContentXml("one").Descendants(), e => (string?)e.Attribute("id") == "keep");
        Assert.Equal("Original", Element(book, "one", "target").Value);
    }
    [Fact]
    public void SvgXmlIdentifiersCanBeEditedWithoutReplacingTheRoot() {
        var book = Book();
        book.AddResource("diagram", "EPUB/diagram.svg", "image/svg+xml", System.Text.Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg'><text xml:id='label'>Before</text></svg>"));
        var expected = book.GetContentXml("diagram").Root!.Elements().Single();
        var replacement = new XElement(expected); replacement.Value = "After";
        book.ApplyContentEdits(new[] { new EpubContentEdit("diagram", "label", expected, replacement) });
        Assert.Equal("After", book.GetContentXml("diagram").Root!.Value);
    }

    private static XElement Element(EpubPublication book, string resource, string id) => book.GetContentXml(resource).Descendants().Single(e => (string?)e.Attribute("id") == id);
    private static EpubPublication Book() {
        var book = EpubPublication.Create("Scoped editing", "en");
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<section id='section'><p id='target'>Original</p><p id='keep'>Keep</p></section>");
        book.AddChapter("two", "EPUB/two.xhtml", "Two", "<p><a id='link' href='one.xhtml#target'>Read</a></p>");
        return book;
    }
}
