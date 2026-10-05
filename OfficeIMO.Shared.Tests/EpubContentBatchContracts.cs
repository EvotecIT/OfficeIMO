using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubContentBatchContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public void TargetAndIncomingLinkCanChangeTogetherWithoutMutatingCallerXml() {
        var book = Book();
        var target = book.GetContentXml("target");
        var source = book.GetContentXml("source");
        target.Descendants(Html + "p").Single().SetAttributeValue("id", "new");
        source.Descendants(Html + "a").Single().SetAttributeValue("href", "target.xhtml#new");
        book.SetContentXml(new Dictionary<string, XDocument> { ["target"] = target, ["source"] = source });
        target.Descendants(Html + "p").Single().Value = "Caller changed";
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal("Original", reopened.GetContentXml("target").Descendants(Html + "p").Single().Value);
        Assert.Equal("target.xhtml#new", (string?)reopened.GetContentXml("source").Descendants(Html + "a").Single().Attribute("href"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DanglingIncomingReferenceOrDuplicateIdentifierRollsBackEveryDocument(bool duplicate) {
        var book = Book();
        byte[] before = book.Write().Bytes;
        var target = book.GetContentXml("target");
        var source = book.GetContentXml("source");
        source.Descendants(Html + "a").Single().Value = "Edited";
        if (duplicate) target.Root!.Element(Html + "body")!.Add(new XElement(Html + "p", new XAttribute("id", "old"), "Duplicate"));
        else target.Descendants(Html + "p").Single().SetAttributeValue("id", "new");
        Assert.Throws<InvalidDataException>(() => book.SetContentXml(new Dictionary<string, XDocument> { ["source"] = source, ["target"] = target }));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void CancellationLeavesRetainedPublicationUnchanged() {
        var book = Book(); byte[] before = book.Write().Bytes;
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => book.SetContentXml(new Dictionary<string, XDocument> { ["target"] = book.GetContentXml("target") }, cancellation.Token));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Editorial", "en");
        book.AddChapter("target", "EPUB/target.xhtml", "Target", "<p id='old'>Original</p>");
        book.AddChapter("source", "EPUB/source.xhtml", "Source", "<p><a href='target.xhtml#old'>Target</a></p>");
        return book;
    }
}
