using System.Xml.Linq;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubManuscriptBoundaryContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public void ForwardAndTransitiveReferencesRetainOrderAndUnaffectedLinks() {
        var result = Import("<h1 id='one'>One</h1><p aria-details='two'>First</p>" +
            "<h1 id='two'>Two</h1><p aria-describedby='three'>Second</p>" +
            "<h1 id='three'>Three</h1><p><a href='#four'>Next</a></p>" +
            "<h1 id='four'>Four</h1><p><a href='#two'>Back</a></p>");
        var book = Reopen(result);
        Assert.Equal(2, book.Spine.Count);
        XDocument first = book.GetContentXml("chapter-1");
        Assert.Equal(new[] { "One", "Two", "Three" }, first.Descendants(Html + "h1").Select(e => e.Value));
        Assert.Equal("chapter-0002.xhtml#four", (string?)first.Descendants(Html + "a").Single().Attribute("href"));
        Assert.Equal("chapter-0001.xhtml#two", (string?)book.GetContentXml("chapter-2").Descendants(Html + "a").Single().Attribute("href"));
        Assert.Equal(new[] { "Two", "Three" }, book.Read().TableOfContents[0].Children.Select(e => e.Label));
    }

    [Theory]
    [InlineData("<section aria-labelledby='label'><h1 id='label'>One</h1><p>First</p><h1>Two</h1><p>Second</p></section>")]
    [InlineData("<section id='group'><h1>One</h1><p aria-details='group'>First</p><h1>Two</h1><p>Second</p></section>")]
    public void AncestorRelationshipsKeepTheWholeOriginalContainer(string content) {
        var book = Reopen(Import(content + "<h1>Independent</h1><p>Third</p>"));
        Assert.Equal(2, book.Spine.Count);
        XElement section = Assert.Single(book.GetContentXml("chapter-1").Descendants(Html + "section"));
        Assert.Equal(new[] { "One", "Two" }, section.Descendants(Html + "h1").Select(e => e.Value));
        Assert.Empty(book.GetContentXml("chapter-2").Descendants(Html + "section"));
    }

    [Fact]
    public void BodyRelationshipsKeepAllContentInOneDocument() {
        var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse(
            "<html lang='en'><head><title>Book</title></head><body aria-describedby='description'>" +
            "<h1>One</h1><p id='description'>Description</p><h1>Two</h1><p>Second</p></body></html>"));
        var book = Reopen(result);
        Assert.Single(book.Spine);
        Assert.Equal("description", (string?)book.GetContentXml("chapter-1").Descendants(Html + "body").Single().Attribute("aria-describedby"));
    }

    [Fact]
    public void EmptyHeadingsAndEmptyXmlTargetsCannotReintroduceASuppressedBoundary() {
        var book = Reopen(Import("<h1>One</h1><p aria-describedby='epub-heading-1'>First</p>" +
            "<h1></h1><h1>Two</h1><div xml:id='epub-heading-1'></div>" +
            "<h1>Independent</h1><p>Third</p>"));
        Assert.Equal(2, book.Spine.Count);
        XDocument first = book.GetContentXml("chapter-1");
        Assert.Equal(new[] { "One", "", "Two" }, first.Descendants(Html + "h1").Select(e => e.Value));
        Assert.Equal("epub-heading-1", (string?)Assert.Single(first.Descendants(Html + "div")).Attribute("id"));
        Assert.DoesNotContain(first.Descendants(Html + "h1"), e => (string?)e.Attribute("id") == "epub-heading-1");
    }

    [Theory]
    [InlineData("<p itemscope='' itemref='target'>First</p>", "<p id='target' itemprop='name'>Target</p>")]
    [InlineData("<svg xmlns='http://www.w3.org/2000/svg' role='img' aria-labelledby='target' viewBox='0 0 10 10'><rect width='10' height='10'/></svg>", "<p id='target'>A square</p>")]
    [InlineData("<img alt='Map' usemap='#target' src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII='>", "<map name='target'><area alt='Destination' href='#last' shape='rect' coords='0,0,1,1'></map>")]
    public void SupportedRelationshipKindsShareBoundaryPlanning(string source, string target) {
        var book = Reopen(Import("<h1>One</h1>" + source + "<h1>Two</h1>" + target + "<h1 id='last'>Independent</h1><p>Third</p>"));
        Assert.Equal(2, book.Spine.Count);
        Assert.Equal(2, book.GetContentXml("chapter-1").Descendants(Html + "h1").Count());
    }

    [Fact]
    public void MissingTargetsStillFailAndLocalReferencesDoNotCollapseOtherChapters() {
        var missing = Import("<h1>One</h1><p aria-describedby='missing'>First</p><h1>Two</h1><p>Second</p>");
        Assert.False(missing.Succeeded);
        Assert.Contains(missing.Report.FidelityDiagnostics, d => d.Code == "EPUB_IMPORT_ID_REFERENCE_INVALID");
        Assert.Throws<InvalidDataException>(() => missing.Publication.Write());
        var local = Import("<h1 id='one'>One</h1><p aria-labelledby='one'>First</p><h1>Two</h1><p>Second</p>");
        Assert.Equal(2, Reopen(local).Spine.Count);
        Assert.DoesNotContain(local.Report.FidelityDiagnostics, d => d.Code == "EPUB_IMPORT_CHAPTER_BOUNDARY_PRESERVED");
    }

    private static EpubManuscriptResult Import(string body) => EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Book</title>" + body));
    private static EpubPublication Reopen(EpubManuscriptResult result) => EpubPublication.Load(new MemoryStream(result.RequireNoLoss().Write().Bytes));
}
