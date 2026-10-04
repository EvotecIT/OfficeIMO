using OfficeIMO.Epub;
using OfficeIMO.Html;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubPublishingResourceContracts {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task RecursiveResourceDuplicatesRespectCapacityAndRejectGenuineExcess(bool svg, bool excess) {
        var requested = new List<string>();
        const string png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
        string parent = svg ? "parent.svg" : "parent.css";
        string content = svg
            ? "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1 1'><image width='1' height='1' href='dot.png'/><image width='1' height='1' href='dot.png'/>" + (excess ? "<image width='1' height='1' href='extra.png'/>" : "") + "</svg>"
            : "@import 'parent.css';p{background:url('dot.png')}h1{background:url('dot.png')}" + (excess ? "h2{background:url('extra.png')}" : "");
        var source = HtmlConversionDocument.Parse("<title>Resources</title>" +
            (svg ? "<img alt='Diagram' src='parent.svg'>" : "<link rel='stylesheet' href='parent.css'>") + "<h1>Resources</h1>",
            new HtmlConversionDocumentOptions { BaseUri = new Uri("https://example.test/book.html") });
        var result = await EpubManuscript.ImportHtmlAsync(source, new EpubManuscriptOptions {
            MaxResourceCount = 2,
            ResourceResolver = (request, _) => {
                requested.Add(request.Uri.AbsolutePath);
                return Task.FromResult<HtmlResolvedResource?>(request.Uri.AbsolutePath == "/" + parent
                    ? new HtmlResolvedResource(System.Text.Encoding.UTF8.GetBytes(content), svg ? "image/svg+xml" : "text/css")
                    : new HtmlResolvedResource(Convert.FromBase64String(png), "image/png"));
            }
        });
        Assert.Equal(new[] { "/" + parent, "/dot.png" }, requested);
        if (excess) {
            Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded && item.LossKind == OfficeConversionLossKind.Failure);
            Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        } else {
            result.Report.RequireNoLoss();
            result.Publication.Write().Report.RequireNoLoss();
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(256)]
    public async Task RepeatedImageCarriersDoNotConsumeDistinctResourceCapacity(int count) {
        var requested = new HashSet<string>();
        var source = HtmlConversionDocument.Parse("<title>Images</title><h1>Images</h1>" +
            string.Concat(Enumerable.Range(0, count).Select(index => "<img alt='A dot' src='image-" + index + ".png'>")),
            new HtmlConversionDocumentOptions { BaseUri = new Uri("https://example.test/book.html") });
        var result = await EpubManuscript.ImportHtmlAsync(source, new EpubManuscriptOptions {
            MaxResourceCount = count,
            ResourceResolver = (request, _) => {
                requested.Add(request.Uri.AbsolutePath);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="), "image/png"));
            }
        });
        result.Report.RequireNoLoss();
        Assert.Equal(count, requested.Count);
        Assert.Equal(count, result.Publication.GetContentXml("chapter-1").Descendants(System.Xml.Linq.XName.Get("img", "http://www.w3.org/1999/xhtml")).Count());
        result.Publication.Write().Report.RequireNoLoss();
    }
    [Theory]
    [InlineData("<map id='sites' name='sites'><area href='#two' alt='Next' shape='rect' coords='0,0,1,1'/></map>")]
    [InlineData("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><a href='#two'><text x='0' y='1'>Next</text></a></svg>")]
    [InlineData("<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink' viewBox='0 0 10 10'><a xlink:href='#two'><text x='0' y='1'>Next</text></a></svg>")]
    public void RetainedAreaAndSvgLinksFollowTheirTargetsAcrossSplitChapters(string link) {
        var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Links</title><h1 id='one'>One</h1>" + link + "<h1 id='two'>Two</h1><p>Target</p>"));
        result.Report.RequireNoLoss();
        var target = result.Publication.GetContentXml("chapter-1").Descendants().Single(element => element.Name.LocalName is "area" or "a");
        Assert.Equal("chapter-0002.xhtml#two", target.Attributes().Single(attribute => attribute.Name.LocalName == "href").Value);
        result.Publication.Write().Report.RequireNoLoss();
    }
    [Fact]
    public async Task ResponsivePictureRetainsBothSelectedSourceAndFallbackImage() {
        var requested = new HashSet<string>();
        var source = HtmlConversionDocument.Parse("<title>Picture</title><h1>Picture</h1><picture><source media='(min-width:1px)' srcset='wide.png 1x'><img alt='A dot' src='fallback.png' srcset='fallback-large.png 2x'></picture>",
            new HtmlConversionDocumentOptions { BaseUri = new Uri("https://example.test/book.html") });
        var result = await EpubManuscript.ImportHtmlAsync(source, new EpubManuscriptOptions {
            ResourceResolver = (request, _) => {
                requested.Add(request.Uri.AbsolutePath);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="), "image/png"));
            }
        });
        result.Report.RequireNoLoss();
        Assert.Contains("/wide.png", requested);
        Assert.Contains("/fallback.png", requested);
        Assert.Contains("/fallback-large.png", requested);
        var picture = result.Publication.GetContentXml("chapter-1").Descendants(System.Xml.Linq.XName.Get("picture", "http://www.w3.org/1999/xhtml")).Single();
        Assert.Contains("resource-", picture.Element(System.Xml.Linq.XName.Get("img", "http://www.w3.org/1999/xhtml"))!.Attribute("src")!.Value);
        result.Publication.Write().Report.RequireNoLoss();
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SharedResolverObjectsRetainEachResourceBaseAndItsDistinctImage(bool svg) {
        string references = svg ? "<img alt='First diagram' src='one/parent.svg'><img alt='Second diagram' src='two/parent.svg'>" : "<link rel='stylesheet' href='one/parent.css'><link rel='stylesheet' href='two/parent.css'>";
        var source = HtmlConversionDocument.Parse("<title>Resources</title>" + references + "<h1>One</h1><p>Text</p>",
            new HtmlConversionDocumentOptions { BaseUri = new Uri("https://example.test/book.html") });
        var shared = new HtmlResolvedResource(System.Text.Encoding.UTF8.GetBytes(svg
            ? "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1 1'><image width='1' height='1' href='dot.svg'/></svg>"
            : "p{background:url('dot.svg')}"), svg ? "image/svg+xml" : "text/css");
        var result = await EpubManuscript.ImportHtmlAsync(source, new EpubManuscriptOptions {
            ResourceResolver = (request, _) => Task.FromResult<HtmlResolvedResource?>(request.Uri.AbsolutePath.Contains("/parent.") ? shared :
                new HtmlResolvedResource(System.Text.Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1 1'><rect width='1' height='1' fill='" + (request.Uri.AbsolutePath.StartsWith("/one/") ? "red" : "blue") + "'/></svg>"), "image/svg+xml"))
        });
        result.Report.RequireNoLoss();
        var parents = result.Publication.Manifest.Where(item => item.Id.StartsWith("manuscript-resource-") &&
            (svg ? item.MediaType == "image/svg+xml" && System.Text.Encoding.UTF8.GetString(result.Publication.GetResourceBytes(item.Id)).Contains("<image") : item.MediaType == "text/css")).ToArray();
        Assert.Equal(2, parents.Length);
        foreach (var parent in parents) {
            string content = System.Text.Encoding.UTF8.GetString(result.Publication.GetResourceBytes(parent.Id));
            var image = result.Publication.Manifest.Single(item => item.MediaType == "image/svg+xml" && item != parent && content.Contains(System.IO.Path.GetFileName(item.Reference.ContainerPath!)));
            string color = parent == parents[0] ? "red" : "blue";
            Assert.Contains("fill=\"" + color + "\"", System.Text.Encoding.UTF8.GetString(result.Publication.GetResourceBytes(image.Id)));
        }
        result.Publication.Write().Report.RequireNoLoss();
    }
    [Fact]
    public void ExcessiveImportFindingsStayBoundedAndCannotHideBehindSuccessfulTruncation() {
        var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Book</title><h1>One</h1>" + string.Concat(Enumerable.Repeat("<p onclick='execute()'>Text</p>", 10_001))));
        Assert.Equal(10_000, result.Report.FidelityDiagnostics.Count);
        Assert.False(result.Succeeded);
        Assert.Equal("EPUB_IMPORT_DIAGNOSTIC_LIMIT", result.Report.FidelityDiagnostics.Last().Code);
    }
    [Fact]
    public void AuthoredPackageAndNavigationCarryLanguageAndMatchingAccessibilityRoles() {
        var publication = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<html lang='pl'><title>Book</title><h1 id='one'>One</h1></html>")).RequireNoLoss();
        using var stream = new MemoryStream(publication.Write().Bytes);
        using var zip = new System.IO.Compression.ZipArchive(stream);
        using var package = zip.GetEntry("EPUB/package.opf")!.Open();
        Assert.Equal("pl", (string?)System.Xml.Linq.XDocument.Load(package).Root!.Attribute(System.Xml.Linq.XNamespace.Xml + "lang"));
        var navigation = publication.GetContentXml("navigation");
        Assert.Contains(navigation.Descendants(System.Xml.Linq.XName.Get("nav", "http://www.w3.org/1999/xhtml")), item => (string?)item.Attribute("role") == "doc-toc");
        var created = EpubPublication.Create("Native book", "en");
        created.AddChapter("one", "EPUB/one.xhtml", "One", "<h1>One</h1>");
        Assert.Equal("doc-toc", (string?)created.GetContentXml("navigation").Descendants(System.Xml.Linq.XName.Get("nav", "http://www.w3.org/1999/xhtml")).Single().Attribute("role"));
    }
    [Fact]
    public void ImageNamesAndDecorativeIntentAreProjectedWithoutInventingAlternativeText() {
        const string image = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
        var imported = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Images</title><h1>Images</h1><img aria-label='A diagram' src='" + image + "'><img role='presentation' src='" + image + "'>"));
        imported.Report.RequireNoLoss();
        var images = imported.Publication.GetContentXml("chapter-1").Descendants(System.Xml.Linq.XName.Get("img", "http://www.w3.org/1999/xhtml")).ToArray();
        Assert.Equal("A diagram", (string?)images[0].Attribute("alt"));
        Assert.Equal(string.Empty, (string?)images[1].Attribute("alt"));
        var unnamed = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Images</title><h1>Images</h1><img src='" + image + "'>"));
        Assert.False(unnamed.Succeeded);
        Assert.Contains(unnamed.Report.FidelityDiagnostics, item => item.Code == "EPUB_IMPORT_IMAGE_ALT_MISSING" && item.LossKind == OfficeConversionLossKind.Failure);
    }
    [Fact]
    public async Task ImportedSvgDependenciesUseTheirOwnBaseAndRemainDeclared() {
        const string png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
        var source = HtmlConversionDocument.Parse("<title>Diagram</title><h1>Diagram</h1><img alt='Diagram with a dot' src='figures/diagram.svg'>",
            new HtmlConversionDocumentOptions { BaseUri = new Uri("https://example.test/book/manuscript.html") });
        var result = await EpubManuscript.ImportHtmlAsync(source, new EpubManuscriptOptions {
            ResourceResolver = (request, _) => Task.FromResult<HtmlResolvedResource?>(request.Uri.AbsolutePath.EndsWith("diagram.svg", StringComparison.Ordinal)
                ? new HtmlResolvedResource(System.Text.Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink' viewBox='0 0 10 10'><image width='10' height='10' xlink:href='../images/dot.png'/></svg>"), "image/svg+xml")
                : request.Uri.AbsolutePath == "/book/images/dot.png" ? new HtmlResolvedResource(Convert.FromBase64String(png), "image/png") : null)
        });
        result.Report.RequireNoLoss();
        Assert.Single(result.Publication.Manifest, item => item.MediaType == "image/png");
        var svg = Assert.Single(result.Publication.Manifest, item => item.MediaType == "image/svg+xml");
        Assert.DoesNotContain("../images/dot.png", System.Text.Encoding.UTF8.GetString(result.Publication.GetResourceBytes(svg.Id)));
        result.Publication.Write().Report.RequireNoLoss();
    }
    [Fact]
    public void AbsoluteAndRelativeSourceDocumentAnchorsBecomeBookLinks() {
        var source = HtmlConversionDocument.Parse("<title>Links</title><h1 id='one'>One</h1><a href='manuscript.html#two'>Relative</a><a href='https://example.test/manuscript.html#two'>Absolute</a><h1 id='two'>Two</h1>",
            new HtmlConversionDocumentOptions { BaseUri = new Uri("https://example.test/manuscript.html") });
        var result = EpubManuscript.ImportHtml(source);
        result.Report.RequireNoLoss();
        var links = result.Publication.GetContentXml("chapter-1").Descendants(System.Xml.Linq.XName.Get("a", "http://www.w3.org/1999/xhtml"));
        Assert.All(links, link => Assert.Equal("chapter-0002.xhtml#two", (string?)link.Attribute("href")));
    }
    [Theory]
    [InlineData("@import 'missing.css';")]
    [InlineData("@font-face{font-family:Test;src:url('missing.woff2')}")]
    [InlineData("@media print{p{background:image-set('missing.png' 1x)}}")]
    public void AuthoredCssMustDeclareEveryDependencyIncludingInactiveResources(string css) {
        var book = EpubPublication.Create("Book", "en");
        book.AddStylesheet("css", "EPUB/styles/main.css", css);
        book.AddChapter("chapter", "EPUB/text/chapter.xhtml", "Chapter", "<p>Body</p>", ["css"]);
        Assert.Throws<InvalidDataException>(() => book.Write());
    }

    [Fact]
    public void CssImportsAreRecursiveAndCyclesAreBoundedWithoutDiscardingContent() {
        var book = EpubPublication.Create("Book", "en");
        book.AddStylesheet("a", "EPUB/styles/a.css", "@import 'b.css'; p{color:red}");
        book.AddStylesheet("b", "EPUB/styles/b.css", "@import 'a.css'; p{color:blue}");
        book.AddChapter("chapter", "EPUB/text/chapter.xhtml", "Chapter", "<p>Body</p>", ["a"]);
        byte[] bytes = book.Write().Bytes;
        using var stream = new MemoryStream(bytes);
        var loaded = EpubPublication.Load(stream);
        loaded.UpdateResource("b", System.Text.Encoding.UTF8.GetBytes("@import 'a.css';p{background:url('missing.png')}"));
        Assert.Throws<InvalidDataException>(() => loaded.Write());
    }

    [Fact]
    public void RemovingAnAssetCannotLeaveRetainedCssWithBrokenDependencies() {
        var book = EpubPublication.Create("Book", "en");
        book.AddResource("image", "EPUB/styles/dot.png", "image/png", [1]);
        book.AddStylesheet("css", "EPUB/styles/main.css", "p{background:url('dot.png')}");
        book.AddChapter("chapter", "EPUB/text/chapter.xhtml", "Chapter", "<p>Body</p>", ["css"]);
        using var stream = new MemoryStream(book.Write().Bytes);
        var loaded = EpubPublication.Load(stream);
        loaded.RemoveResource("image");
        Assert.Throws<InvalidDataException>(() => loaded.Write());
    }

    [Fact]
    public void TaskListStatesAndHtmlMetadataAreRetainedAsStaticBookContent() {
        var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Tasks</title><meta name='author' content='Writer'><meta name='description' content='Task notes'><h1>Tasks</h1><ul><li><input type='checkbox' disabled checked>Done</li><li><input type='checkbox' disabled>Open</li></ul>"));
        result.Report.RequireNoLoss();
        Assert.Equal("Writer", result.Publication.Creator);
        string content = result.Publication.GetContentXml("chapter-1").ToString();
        Assert.Contains("[x] ", content); Assert.Contains("[ ] ", content); Assert.DoesNotContain("<input", content);
        result.Publication.Write().Report.RequireNoLoss();
    }
}
