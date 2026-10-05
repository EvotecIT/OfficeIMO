using OfficeIMO.Epub;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubGlossaryAuthoringContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly XNamespace Ops = "http://www.idpf.org/2007/ops";

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void GlossarySupportsMultipleOccurrencesAndOptionalReturnLinks(bool sameDocument, bool backlinks) {
        var book = Book(sameDocument);
        string glossaryId = sameDocument ? "source" : "glossary";
        book.AddGlossaryEntry(glossaryId, "terms", "term-one", "A & B", "<p>A <em>formatted</em> definition.</p>");
        book.AddGlossaryEntry(glossaryId, "terms", "term-two", "Second", "<p>Another definition.</p>");
        book.LinkGlossaryTerm("source", "ref-one", glossaryId, "term-one", backlinks ? "Return to first occurrence" : null);
        book.LinkGlossaryTerm("source", "ref-two", glossaryId, "term-one", backlinks ? "Return to second occurrence" : null);
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        XDocument glossary = reopened.GetContentXml(glossaryId);
        XElement term = ById(glossary, "term-one");
        Assert.Equal(Html + "dt", term.Name);
        Assert.Equal("A & B", term.Element(Html + "dfn")!.Value);
        Assert.Equal("glossary", (string?)term.Parent!.Parent!.Attribute(Ops + "type"));
        Assert.Equal("doc-glossary", (string?)term.Parent.Parent!.Attribute("role"));
        Assert.Equal("Retained heading", term.Parent.Parent.Element(Html + "h2")!.Value);
        Assert.Equal(new[] { "term-one", "term-two" }, term.Parent.Elements(Html + "dt").Select(e => (string?)e.Attribute("id")));
        Assert.Equal("formatted", term.ElementsAfterSelf().First().Descendants(Html + "em").Single().Value);
        XElement reference = ById(reopened.GetContentXml("source"), "ref-one");
        Assert.Equal("retained", (string?)reference.Attribute("class"));
        Assert.Equal("doc-glossref", (string?)reference.Attribute("role"));
        Assert.Equal("glossref", (string?)reference.Attribute(Ops + "type"));
        var forward = EpubReference.Resolve("EPUB/text/source.xhtml", reference.Attribute("href")!.Value);
        Assert.Equal(sameDocument ? "EPUB/text/source.xhtml" : "EPUB/back/glossary.xhtml", forward.ContainerPath);
        Assert.Equal("term-one", forward.Fragment);
        var returns = glossary.Descendants(Html + "a").Where(e => (string?)e.Attribute("role") == "doc-backlink").ToArray();
        Assert.Equal(backlinks ? 2 : 0, returns.Length);
        if (backlinks) {
            Assert.Equal(new[] { "ref-one", "ref-two" }, returns.Select(e => EpubReference.Resolve(forward.ContainerPath!, e.Attribute("href")!.Value).Fragment));
            Assert.All(returns, e => Assert.Equal("EPUB/text/source.xhtml", EpubReference.Resolve(forward.ContainerPath!, e.Attribute("href")!.Value).ContainerPath));
        }
    }

    [Fact]
    public void LinksUseBothHtmlBasesAndPreserveUntouchedPayloads() {
        var book = Book(false);
        book.AddGlossaryEntry("glossary", "terms", "term-one", "Term", "<p>Definition.</p>");
        foreach (string id in new[] { "source", "glossary" }) {
            var content = book.GetContentXml(id);
            content.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", "../assets/")));
            book.SetContentXml(id, content);
        }
        byte[] untouched = book.GetResourceBytes("glossary");
        book.LinkGlossaryTerm("source", "ref-one", "glossary", "term-one");
        Assert.Equal(untouched, book.GetResourceBytes("glossary"));
        book.LinkGlossaryTerm("source", "ref-two", "glossary", "term-one", "Return");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        string href = ById(reopened.GetContentXml("source"), "ref-two").Attribute("href")!.Value;
        Assert.Equal("EPUB/back/glossary.xhtml", EpubReference.Resolve("EPUB/text/source.xhtml", "../assets/", href).ContainerPath);
        string back = reopened.GetContentXml("glossary").Descendants(Html + "a").Single().Attribute("href")!.Value;
        Assert.Equal("EPUB/text/source.xhtml", EpubReference.Resolve("EPUB/back/glossary.xhtml", "../assets/", back).ContainerPath);
    }

    [Theory]
    [InlineData("<p id='term-one'>Duplicate.</p>")]
    [InlineData("<p aria-describedby='missing'>Dangling IDREF.</p>")]
    public void RejectedDefinitionDoesNotChangeGlossaryOrItsSemantics(string body) {
        var book = Book(false);
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.AddGlossaryEntry("glossary", "terms", "term-one", "Term", body));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData("linked")]
    [InlineData("missing")]
    [InlineData("role")]
    [InlineData("budget")]
    [InlineData("cancel")]
    public void RejectedLinkLeavesBothDocumentsUnchanged(string failure) {
        var book = Book(false);
        book.AddGlossaryEntry("glossary", "terms", "term-one", "Term", "<p>Definition.</p>");
        if (failure == "linked" || failure == "role") {
            var source = book.GetContentXml("source");
            ById(source, "ref-one").SetAttributeValue(failure == "linked" ? "href" : "role", failure == "linked" ? "#ref-one" : "button");
            book.SetContentXml("source", source);
        }
        byte[] before = book.Write().Bytes;
        if (failure == "budget") {
            using var zip = new ZipArchive(new MemoryStream(before));
            book = EpubPublication.Load(new MemoryStream(before), new EpubPublicationLoadOptions { MaxExpandedBytes = zip.Entries.Sum(e => e.Length) + 32 });
        }
        using var cancellation = new CancellationTokenSource();
        if (failure == "cancel") cancellation.Cancel();
        Action operation = () => book.LinkGlossaryTerm("source", "ref-one", "glossary", failure == "missing" ? "absent" : "term-one", "Return", cancellation.Token);
        if (failure == "cancel") Assert.Throws<OperationCanceledException>(operation);
        else Assert.Throws<InvalidDataException>(operation);
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void InvalidContainerAndScriptedDefinitionAreRejectedAtomically() {
        var book = Book(false);
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.AddGlossaryEntry("glossary", "section", "term", "Term", "<p>Definition.</p>"));
        Assert.Throws<NotSupportedException>(() => book.AddGlossaryEntry("glossary", "terms", "term", "Term", "<script>bad()</script>"));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static XElement ById(XDocument document, string id) => document.Descendants().Single(e => (string?)e.Attribute("id") == id);
    private static EpubPublication Book(bool sameDocument) {
        var book = EpubPublication.Create("Glossary", "en");
        string glossary = "<section id='section' aria-labelledby='glossary-heading'><h2 id='glossary-heading'>Retained heading</h2><dl id='terms'/></section>";
        book.AddChapter("source", "EPUB/text/source.xhtml", "Chapter", "<h1>Chapter</h1><p><a id='ref-one' class='retained'>Term</a> and <a id='ref-two'>the same term</a>.</p>" + (sameDocument ? glossary : ""));
        if (!sameDocument) book.AddChapter("glossary", "EPUB/back/glossary.xhtml", "Glossary", "<h1>Glossary</h1>" + glossary);
        return book;
    }
}
