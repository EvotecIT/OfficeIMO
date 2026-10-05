using OfficeIMO.Epub;
using OfficeIMO.Html;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubPublishingPreflightContracts {
    [Fact]
    public void PreflightReportsNativeFailuresAndNeverClaimsExternalQualification() {
        var book = EpubPublication.Create("Book", "en");
        book.AddChapter("c", "EPUB/c.xhtml", "Chapter", "<h1 id='x'>First</h1><p id='x'>Duplicate</p>");
        var invalid = book.Preflight();
        Assert.True(invalid.HasErrors);
        Assert.Contains(invalid.Checks, check => check.Code == "native-save" && check.Status == EpubPreflightStatus.Failed);
        Assert.Contains(invalid.Checks, check => check.Code == "content-identifiers" && check.Diagnostics.Any(item => item.Path == "EPUB/c.xhtml"));
        XDocument content = book.GetContentXml("c");
        content.Descendants().Single(element => element.Name.LocalName == "p").Attribute("id")!.Remove();
        book.SetContentXml("c", content);
        byte[] before = book.Write().Bytes;
        var valid = book.Preflight();
        Assert.False(valid.HasErrors);
        Assert.True(valid.HasUncheckedItems);
        Assert.Equal(3, valid.Checks.Count(check => check.Status == EpubPreflightStatus.NotChecked));
        Assert.Equal(before, book.Write().Bytes);
        using var cancelled = new System.Threading.CancellationTokenSource();
        cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => book.Preflight(cancellationToken: cancelled.Token));
    }

    [Theory]
    [InlineData("aria-describedby")]
    [InlineData("aria-labelledby")]
    [InlineData("aria-details")]
    public void SplitCannotSilentlyBreakDocumentLocalRelationships(string attribute) {
        var html = HtmlConversionDocument.Parse("<title>Book</title><h1>First</h1><p id='description'>Description</p>" +
            "<h1>Second</h1><section " + attribute + "='description'><p>Content</p></section>");
        var split = EpubManuscript.ImportHtml(html);
        Assert.False(split.Succeeded);
        Assert.Contains(split.Report.FidelityDiagnostics, item => item.Code == "EPUB_IMPORT_ID_REFERENCE_INVALID");
        Assert.Throws<InvalidOperationException>(() => split.RequireNoLoss());
        Assert.Throws<InvalidDataException>(() => split.Publication.Write());

        var whole = EpubManuscript.ImportHtml(html, new EpubManuscriptOptions { ChapterHeadingLevel = 0 });
        byte[] bytes = whole.RequireNoLoss().Write().Bytes;
        var reopened = EpubPublication.Load(new MemoryStream(bytes));
        XDocument content = reopened.GetContentXml("chapter-1");
        Assert.Contains(content.Descendants().Attributes(attribute), value => value.Value == "description");
        Assert.Contains(content.Descendants().Attributes("id"), value => value.Value == "description");
    }

    [Fact]
    public void DuplicateIdsAreRejectedWithinADocumentButAllowedAcrossChapters() {
        var book = EpubPublication.Create("Book", "en");
        book.AddChapter("first", "EPUB/first.xhtml", "First", "<h1 id='same'>First</h1><p id='same'>Duplicate</p>");
        Assert.Throws<InvalidDataException>(() => book.Write());
        XDocument content = book.GetContentXml("first");
        content.Descendants().Single(element => element.Name.LocalName == "p").SetAttributeValue("id", "different");
        book.SetContentXml("first", content);
        book.AddChapter("second", "EPUB/second.xhtml", "Second", "<h1 id='same'>Second</h1>");
        Assert.NotEmpty(book.Write().Bytes);
    }

    [Fact]
    public void UneditedInvalidSourceIsPreservedButPreflightReportsIt() {
        byte[] bytes = EpubIntegrityFixtures.OneChapter("<p id='same'>One</p><p id='same'>Two</p>");
        var book = EpubPublication.Load(new MemoryStream(bytes));
        Assert.Equal(bytes, book.Write().Bytes);
        Assert.Contains(book.Preflight().Checks, check => check.Code == "content-identifiers" && check.Status == EpubPreflightStatus.Failed);
        string id = book.Spine[0].ManifestId;
        XDocument content = book.GetContentXml(id);
        content.Root!.Add(new XComment("An edit"));
        book.SetContentXml(id, content);
        Assert.Throws<InvalidDataException>(() => book.Write());
    }

    [Fact]
    public void PreflightChecksImageAlternativesEvenForUneditedImports() {
        var book = EpubPublication.Create("Images", "en");
        book.AddResource("image", "EPUB/image.png", "image/png", Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="));
        book.AddChapter("c", "EPUB/c.xhtml", "Image", "<h1>Image</h1><img src='image.png'/>");
        var imported = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Contains(imported.Preflight().Checks, check => check.Code == "image-alternative-presence" && check.Status == EpubPreflightStatus.Failed);
        XDocument content = imported.GetContentXml("c");
        content.Descendants().Single(element => element.Name.LocalName == "img").SetAttributeValue("alt", "");
        imported.SetContentXml("c", content);
        Assert.Contains(imported.Preflight().Checks, check => check.Code == "image-alternative-presence" && check.Status == EpubPreflightStatus.Passed);
        Assert.True(imported.Preflight().HasUncheckedItems);
    }

    [Fact]
    public void TableHeaderIdListsMustResolveWithinTheDocument() {
        var book = EpubPublication.Create("Table", "en");
        book.AddChapter("c", "EPUB/c.xhtml", "Table", "<table><tr><th id='a'>A</th><th id='b'>B</th></tr><tr><td headers='a b'>Both</td></tr></table>");
        Assert.NotEmpty(book.Write().Bytes);
        XDocument content = book.GetContentXml("c");
        content.Descendants().Attributes("headers").Single().Value = "a missing";
        book.SetContentXml("c", content);
        Assert.Throws<InvalidDataException>(() => book.Write());
    }

    [Fact]
    public void InvalidDublinCoreNameIsRejectedWithoutMutatingMetadata() {
        var book = EpubPublication.Create("Book", "en");
        book.AddChapter("c", "EPUB/c.xhtml", "Chapter", "<h1>Chapter</h1>");
        byte[] before = book.Write().Bytes;
        Assert.Throws<ArgumentException>(() => book.AddDublinCoreMetadata("publsiher", "Wrong"));
        Assert.Equal(before, book.Write().Bytes);
        book.AddDublinCoreMetadata("publisher", "Publisher");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Contains(reopened.GetPackageXml().Descendants(), element =>
            element.Name == XName.Get("publisher", "http://purl.org/dc/elements/1.1/") && element.Value == "Publisher");
    }
}
