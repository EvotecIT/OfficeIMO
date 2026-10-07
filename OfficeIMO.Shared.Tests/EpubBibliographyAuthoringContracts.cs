using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubBibliographyAuthoringContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly XNamespace Ops = "http://www.idpf.org/2007/ops";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BibliographyRetainsFormattingAndSupportsRepeatedCitations(bool sameDocument) {
        var book = Book(sameDocument);
        string target = sameDocument ? "source" : "references";
        book.AddBibliographyEntry(target, "entries", "work", "Writer. <em>A book</em>. 2026.");
        book.LinkBibliographyEntry("source", "cite-one", target, "work", "Return to first citation");
        book.LinkBibliographyEntry("source", "cite-two", target, "work", "Return to second citation");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var bibliography = reopened.GetContentXml(target);
        var entry = ById(bibliography, "work");
        Assert.Equal("A book", entry.Element(Html + "em")!.Value);
        Assert.Null(entry.Attribute("role"));
        Assert.Null(entry.Attribute(Ops + "type"));
        Assert.Equal("doc-bibliography", (string?)entry.Parent!.Parent!.Attribute("role"));
        Assert.Equal("bibliography", (string?)entry.Parent.Parent.Attribute(Ops + "type"));
        var marker = ById(reopened.GetContentXml("source"), "cite-one");
        Assert.Equal("[1]", marker.Value);
        Assert.Equal("doc-biblioref", (string?)marker.Attribute("role"));
        var forward = EpubReference.Resolve("EPUB/text/source.xhtml", marker.Attribute("href")!.Value);
        Assert.Equal(sameDocument ? "EPUB/text/source.xhtml" : "EPUB/back/references.xhtml", forward.ContainerPath);
        Assert.Equal("work", forward.Fragment);
        var backs = entry.Descendants(Html + "a").ToArray();
        Assert.Equal(new[] { "cite-one", "cite-two" }, backs.Select(e => EpubReference.Resolve(forward.ContainerPath!, e.Attribute("href")!.Value).Fragment));
        Assert.All(backs, e => Assert.Equal("doc-backlink", (string?)e.Attribute("role")));
    }

    [Fact]
    public void ForwardOnlyCitationPreservesBibliographyBytesAndRespectsSourceBase() {
        var book = Book(false);
        book.AddBibliographyEntry("references", "entries", "work", "A book.");
        var content = book.GetContentXml("source");
        content.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", "../assets/")));
        book.SetContentXml("source", content);
        byte[] before = book.GetResourceBytes("references");
        book.LinkBibliographyEntry("source", "cite-one", "references", "work");
        Assert.Equal(before, book.GetResourceBytes("references"));
        var marker = ById(book.GetContentXml("source"), "cite-one");
        Assert.Equal("EPUB/back/references.xhtml", EpubReference.Resolve("EPUB/text/source.xhtml", "../assets/", marker.Attribute("href")!.Value).ContainerPath);
        book.Write();
    }

    [Theory]
    [InlineData("duplicate")]
    [InlineData("missing-target")]
    [InlineData("existing-link")]
    public void FailedEditsAreAtomic(string failure) {
        var book = Book(false);
        book.AddBibliographyEntry("references", "entries", "work", "A book.");
        if (failure == "existing-link") book.LinkBibliographyEntry("source", "cite-one", "references", "work");
        byte[] before = book.Write().Bytes;
        Action edit = failure == "duplicate"
            ? () => book.AddBibliographyEntry("references", "entries", "work", "Duplicate.")
            : () => book.LinkBibliographyEntry("source", "cite-one", "references", failure == "missing-target" ? "absent" : "work", "Return");
        Assert.Throws<InvalidDataException>(edit);
        Assert.Equal(before, book.Write().Bytes);
    }

    private static XElement ById(XDocument document, string id) => document.Descendants().Single(e => (string?)e.Attribute("id") == id);
    private static EpubPublication Book(bool sameDocument) {
        var book = EpubPublication.Create("References", "en");
        string references = "<section><h2>References</h2><ol id='entries'/></section>";
        book.AddChapter("source", "EPUB/text/source.xhtml", "Chapter", "<h1>Chapter</h1><p>First <a id='cite-one'>[1]</a>, second <a id='cite-two'>[1]</a>.</p>" + (sameDocument ? references : ""));
        if (!sameDocument) book.AddChapter("references", "EPUB/back/references.xhtml", "References", "<h1>Sources</h1>" + references);
        return book;
    }
}
