using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using System.Threading;

namespace OfficeIMO.Chm.Tests;

public sealed class ConversionTests {
    [Fact]
    public void LinkedBookRetainsTopicAnchorsTablesAndPortableImages() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Book());
        var html = book.ToHtmlDocumentResult();
        Assert.Equal(2, html.Value.Document.QuerySelectorAll("section[data-chm-topic]").Count);
        Assert.NotNull(html.Value.Document.QuerySelector("#chm-topic-2-part"));
        Assert.Equal("#chm-topic-2-part", html.Value.Document.QuerySelector("a")!.GetAttribute("href"));
        Assert.StartsWith("data:image/png;base64,", html.Value.Document.QuerySelector("img")!.GetAttribute("src"));
        var markdown = book.ToMarkdownResult();
        Assert.Contains("Café", markdown.Value); Assert.Contains("Second topic", markdown.Value); Assert.Contains("42", markdown.Value);
        Assert.Contains("#chm-topic-2-part", markdown.Value);
        Assert.Contains(markdown.Report.FidelityDiagnostics, item => item.Code == "CHM_INDEX_METADATA" && item.LossKind == OfficeConversionLossKind.Omission);
        Assert.Throws<OfficeConversionException>(() => markdown.RequireNoLoss());
    }

    [Fact]
    public void BookWideAnchorNamesRetainTableHeadersAndAccessibleRelationships() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/table.html"] = ChmFixture.Html("<p id='description'>Values</p><table aria-describedby='description'><tr><th id='column'>Name</th></tr>" +
                "<tr><td headers='column'>Example</td></tr></table>")
        }));
        var projection = book.ToHtmlDocumentResult();
        Assert.Equal("chm-topic-1-column", projection.Value.Document.QuerySelector("td")!.GetAttribute("headers"));
        Assert.Equal("chm-topic-1-description", projection.Value.Document.QuerySelector("table")!.GetAttribute("aria-describedby"));
        var epub = book.ToEpubPublicationResult(); Assert.True(epub.Succeeded);
        Assert.NotEmpty(epub.Publication.Write().RequireValue());
    }

    [Fact]
    public void RepeatedAdjacentStylesheetsAreCollapsedWithoutReorderingTheCascade() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/1.html"] = ChmFixture.Html("<link rel='stylesheet' href='/a.css'><p>First</p>"),
            ["/2.html"] = ChmFixture.Html("<link rel='stylesheet' href='/a.css'><p>Second</p>"),
            ["/3.html"] = ChmFixture.Html("<link rel='stylesheet' href='/b.css'><link rel='stylesheet' href='/a.css'><p>Third</p>")
        }));
        var projection = book.ToHtmlDocumentResult();
        Assert.Equal(new[] { "chm://archive/a.css", "chm://archive/b.css", "chm://archive/a.css" },
            projection.Value.Document.QuerySelectorAll("link[rel=stylesheet]").Select(link => link.GetAttribute("href")));
    }

    [Fact]
    public async Task EpubReopensWithHierarchyAndEmbeddedResources() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Book());
        EpubManuscriptResult imported = await book.ToEpubPublicationResultAsync();
        Assert.True(imported.Succeeded, string.Join("\n", imported.Report.FidelityDiagnostics.Select(item => item.Code + ": " + item.Message)));
        var result = book.ToEpubBytesResult(writeOptions: new EpubWriteOptions { ModifiedAt = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero) });
        using var stream = new MemoryStream(result.Value); EpubDocument reopened = EpubDocument.Load(stream);
        Assert.Contains(reopened.Chapters, chapter => chapter.Text.Contains("Café"));
        Assert.Contains(reopened.Chapters, chapter => chapter.Text.Contains("42"));
        EpubNavigationItem group = Assert.Single(reopened.TableOfContents); Assert.Equal("Guide", group.Label); Assert.Equal(2, group.Children.Count);
        Assert.EndsWith("-part", group.Children[1].Fragment);
        Assert.Contains(imported.Publication.Manifest, item => item.MediaType == "image/png");
        Assert.DoesNotContain(reopened.Diagnostics, item => item.Severity == EpubDiagnosticSeverity.Error);
    }

    [Fact]
    public void PdfKeepsTopicPageBoundariesAndSearchableContent() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Book());
        var result = book.ToPdfBytesResult();
        PdfReadDocument read = PdfReadDocument.Open(result.Value);
        Assert.Equal(2, read.Pages.Count); string text = read.ExtractText();
        Assert.Contains("Café", text); Assert.Contains("Second topic", text); Assert.Contains("42", text);
        Assert.NotEmpty(read.Pages[0].GetImages());
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "CHM_PDF_TOPIC_LINKS");
    }

    [Fact]
    public async Task ResourceResolverNeverDelegatesToExternalLocations() {
        ChmDocument original = ChmDocument.Load(ChmFixture.Book());
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/topic.html"] = ChmFixture.Html("<img src='https://example.invalid/pixel.png'><img src='file:///etc/passwd'><img src='pixel.png'>"),
            ["/pixel.png"] = original.FindEntry("/pixel.png")!.GetBytes()
        }));
        int delegated = 0;
        var rendering = new HtmlToPdfOptions { ResourceResolver = (request, token) => { delegated++; return Task.FromResult<HtmlResolvedResource?>(null); } };
        rendering.ResourcePolicy.AllowRemoteResourceResolution = true;
        var result = await book.ToPdfDocumentResultAsync(pdfOptions: rendering);
        Assert.Equal(0, delegated); Assert.NotEmpty(result.ToBytes());
        Assert.Single(PdfReadDocument.Open(result.ToBytes()).Pages[0].GetImages());
        rendering.ResourcePolicy.AllowEmbeddedPackageResources = false;
        var withoutImages = await book.ToPdfDocumentResultAsync(pdfOptions: rendering);
        Assert.Empty(PdfReadDocument.Open(withoutImages.ToBytes()).Pages[0].GetImages());
    }

    [Fact]
    public void SelectionAggregateBudgetsAndCancellationAreExplicit() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Book());
        var result = book.ToMarkdownResult(new ChmConversionOptions { TopicPaths = new[] { "/guide/details.html" } });
        Assert.Contains("Second topic", result.Value); Assert.DoesNotContain("Café", result.Value);
        Assert.Single(result.Report.TopicPaths);
        Assert.Throws<ArgumentException>(() => book.ToMarkdownResult(new ChmConversionOptions { TopicPaths = new[] { "/missing.html" } }));
        Assert.Equal("CHM_CONVERSION_LIMIT", Assert.Throws<ChmReadException>(() => book.ToMarkdownResult(new ChmConversionOptions { MaxTotalHtmlCharacters = 10 })).Code);
        Assert.Equal("CHM_CONVERSION_LIMIT", Assert.Throws<ChmReadException>(() => book.ToMarkdownResult(new ChmConversionOptions { MaxTopics = 1 })).Code);
        Assert.ThrowsAny<OperationCanceledException>(() => book.ToPdfDocumentResult(cancellationToken: new CancellationToken(true)));
    }

    [Fact]
    public void ActiveContentIsInertAndItsOmissionIsReported() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/topic.html"] = ChmFixture.Html("<p onclick='alert(1)'>Visible</p><script>danger()</script><object classid='clsid:example'>Control</object>")
        }));
        var projection = book.ToHtmlDocumentResult();
        Assert.DoesNotContain("danger()", projection.Value.SourceHtml); Assert.DoesNotContain("onclick", projection.Value.SourceHtml);
        Assert.Contains(projection.Report.FidelityDiagnostics, item => item.Code == "CHM_ACTIVE_CONTENT_OMITTED");
        Assert.Contains("danger()", book.Topics[0].ReadHtml());
    }
}
