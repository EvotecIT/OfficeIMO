using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubManuscriptContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private const string ImageData = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

    [Fact]
    public void Import_NormalizesLegacyHelpTablesAndNavigationWithoutLosingLinks() {
        var source = HtmlConversionDocument.Parse("<title>Manual</title><h1>Manual</h1><table width='100%' summary='Navigation' border='0' cellpadding='4' cellspacing='0'>" +
            "<tr><td align='right' valign='top' width='40%' style='text-align:left'>Next</td></tr></table>" +
            "<ul type='circle'><li>Option</li></ul><dl><dt><a href='#legacy'>Topic</a><dl><dt>Child</dt></dl></dt></dl>" +
            "<dl><dt>Defined</dt><dd>Definition</dd><dt>Undescribed</dt></dl>" +
            "<a name='legacy' id='modern'>Target</a><hr align='left' width='100'>");
        var result = EpubManuscript.ImportHtml(source);
        Assert.True(result.Succeeded);
        XDocument chapter = result.Publication.GetContentXml("chapter-1");
        XElement table = Assert.Single(chapter.Descendants(Html + "table"));
        Assert.Equal("Navigation", (string?)table.Attribute("aria-description"));
        Assert.Contains("width:100%", (string?)table.Attribute("style"));
        Assert.Contains("border:0px solid", (string?)table.Attribute("style"));
        XElement cell = Assert.Single(chapter.Descendants(Html + "td"));
        string style = (string)cell.Attribute("style")!;
        Assert.Contains("padding:4px", style); Assert.Contains("vertical-align:top", style);
        Assert.True(style.IndexOf("text-align:right", StringComparison.Ordinal) < style.IndexOf("text-align:left", StringComparison.Ordinal));
        Assert.Contains(chapter.Descendants(Html + "ul"), list => (string?)list.Attribute("style") == "list-style-type:circle;");
        XElement definitions = Assert.Single(chapter.Descendants(Html + "dl"));
        Assert.Equal(Html + "dd", definitions.Elements().Last().Name); Assert.Empty(definitions.Elements().Last().Nodes());
        Assert.Equal("Definition", definitions.Elements(Html + "dd").First().Value);
        Assert.Contains(chapter.Descendants(Html + "a"), anchor => (string?)anchor.Attribute("href") == "#legacy");
        Assert.Contains(chapter.Descendants().Attributes("id"), id => id.Value == "legacy");
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "EPUB_IMPORT_LEGACY_LIST_NORMALIZED" && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.NotEmpty(result.Publication.Write().RequireValue());
    }

    [Fact]
    public void Import_CollectsEmbeddedImagesOnceAndRetainsAlternativeText() {
        var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Images</title><h1>First</h1><img alt='A dot' src='" + ImageData + "'><h1>Second</h1><img alt='' src='" + ImageData + "'>"));
        result.Report.RequireNoLoss();
        EpubManifestItem image = Assert.Single(result.Publication.Manifest, item => item.MediaType == "image/png");
        Assert.Equal("A dot", (string?)result.Publication.GetContentXml("chapter-1").Descendants(Html + "img").Single().Attribute("alt"));
        Assert.StartsWith("../resources/", (string?)result.Publication.GetContentXml("chapter-2").Descendants(Html + "img").Single().Attribute("src"));
        Assert.NotEmpty(result.Publication.GetResourceBytes(image.Id));
        result.Publication.Write().Report.RequireNoLoss();
    }

    [Fact]
    public async Task Import_PreservesStylesheetOrderMediaAndBodyClasses() {
        var source = HtmlConversionDocument.Parse("<html class='theme' dir='rtl'><head><title>Styles</title><style>p{color:red}</style><link rel='stylesheet' href='print.css' media='print'><style>p{color:blue}</style></head><body class='manuscript'><h1>Chapter</h1><p>Text</p></body></html>",
            new HtmlConversionDocumentOptions { BaseUri = new Uri("https://example.test/book.html") });
        var result = await EpubManuscript.ImportHtmlAsync(source, new EpubManuscriptOptions {
            ResourceResolver = (_, _) => Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(System.Text.Encoding.UTF8.GetBytes("p{color:green}"), "text/css"))
        });
        result.Report.RequireNoLoss();
        XDocument chapter = result.Publication.GetContentXml("chapter-1");
        XElement[] links = chapter.Descendants(Html + "link").ToArray();
        Assert.Equal("print", (string?)links[2].Attribute("media"));
        Assert.Contains("red", System.Text.Encoding.UTF8.GetString(result.Publication.GetResourceBytes("manuscript-style-1")));
        Assert.Contains("blue", System.Text.Encoding.UTF8.GetString(result.Publication.GetResourceBytes("manuscript-style-2")));
        Assert.Equal("manuscript", (string?)chapter.Descendants(Html + "body").Single().Attribute("class"));
        Assert.Equal("theme", (string?)chapter.Root!.Attribute("class"));
        Assert.Equal("rtl", (string?)chapter.Root!.Attribute("dir"));
        result.Publication.Write().Report.RequireNoLoss();
    }

    [Fact]
    public void Import_RetainsEpubFootnoteVocabularyAndSvgNamespaces() {
        var source = HtmlConversionDocument.Parse("<title>Semantics</title><h1>Chapter</h1><p><a epub:type='noteref' href='#note'>1</a></p><aside epub:type='footnote' id='note'>Note</aside><svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><title>A shape</title><rect width='10' height='10'/></svg>");
        var result = EpubManuscript.ImportHtml(source);
        result.Report.RequireNoLoss();
        XDocument chapter = result.Publication.GetContentXml("chapter-1");
        Assert.Equal("noteref", (string?)chapter.Descendants(Html + "a").Single().Attribute(XName.Get("type", "http://www.idpf.org/2007/ops")));
        Assert.Single(chapter.Descendants(XName.Get("svg", "http://www.w3.org/2000/svg")));
        result.Publication.Write().Report.RequireNoLoss();
    }

    [Fact]
    public async Task Import_CollectsInactiveStylesheetImportsFontsAndImagesAndRewritesTheirCarriers() {
        var source = HtmlConversionDocument.Parse("<title>Resources</title><link rel='stylesheet' href='styles/main.css'><h1>Chapter</h1><p style='background-image:image-set(\"images/dot.png\" 1x)'>Body</p>",
            new HtmlConversionDocumentOptions { BaseUri = new Uri("https://example.test/book/manuscript.html") });
        var requested = new List<string>();
        var result = await EpubManuscript.ImportHtmlAsync(source, new EpubManuscriptOptions {
            ResourceResolver = (request, token) => {
                requested.Add(request.Uri.AbsoluteUri);
                HtmlResolvedResource? resource = request.Uri.AbsolutePath switch {
                    "/book/styles/main.css" => new HtmlResolvedResource(System.Text.Encoding.UTF8.GetBytes("@import 'nested.css' print;@media print{p{background-image:url('../images/dot.png')}}@font-face{font-family:Book;src:url('../fonts/book.woff2')}"), "text/css"),
                    "/book/styles/nested.css" => new HtmlResolvedResource(System.Text.Encoding.UTF8.GetBytes("@supports(display:grid){p{background-image:image-set(\"../images/dot.png\" 1x)}}"), "text/css"),
                    "/book/images/dot.png" => new HtmlResolvedResource(HtmlDataUri.TryParse(ImageData, out var data) ? data.DecodeBytes() : throw new InvalidOperationException(), "image/png"),
                    "/book/fonts/book.woff2" => new HtmlResolvedResource(new byte[] { 1, 2, 3 }, "font/woff2"),
                    _ => null
                };
                return Task.FromResult(resource);
            }
        });
        result.Report.RequireNoLoss();
        Assert.Equal(4, requested.Count);
        Assert.Contains("https://example.test/book/styles/nested.css", requested);
        Assert.Contains("https://example.test/book/fonts/book.woff2", requested);
        Assert.Equal(2, result.Publication.Manifest.Count(item => item.MediaType == "text/css" && item.Id.StartsWith("manuscript-resource-")));
        foreach (EpubManifestItem stylesheet in result.Publication.Manifest.Where(item => item.MediaType == "text/css" && item.Id.StartsWith("manuscript-resource-"))) {
            string css = System.Text.Encoding.UTF8.GetString(result.Publication.GetResourceBytes(stylesheet.Id));
            Assert.DoesNotContain("../images/", css);
            Assert.DoesNotContain("../fonts/", css);
        }
        string mainCss = System.Text.Encoding.UTF8.GetString(result.Publication.GetResourceBytes("manuscript-resource-1"));
        Assert.Contains(" print;", mainCss);
        Assert.Contains("@media print", mainCss);
        Assert.Contains("url(\"resource-", mainCss);
        Assert.Contains("../resources/", (string?)result.Publication.GetContentXml("chapter-1").Descendants(Html + "p").Single().Attribute("style"));
        result.Publication.Write().Report.RequireNoLoss();
    }

    [Theory]
    [InlineData(1)]
    [InlineData(1024)]
    public async Task Import_ReportsMissingOrOverBudgetResourcesWithoutImplicitNetworkAccess(long budget) {
        var source = HtmlConversionDocument.Parse("<title>Missing</title><h1>Chapter</h1><img alt='A dot' src='https://example.test/missing.png'>");
        int calls = 0;
        var result = await EpubManuscript.ImportHtmlAsync(source, new EpubManuscriptOptions {
            MaxResourceBytes = budget, MaxTotalResourceBytes = budget,
            ResourceResolver = (request, token) => { calls++; return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(new byte[2048], "image/png")); }
        });
        Assert.Equal(1, calls);
        Assert.False(result.Report.Succeeded);
        Assert.Throws<InvalidOperationException>(() => result.RequireValue());
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "EPUB_IMPORT_RESOURCE_MISSING");
        Assert.Throws<InvalidOperationException>(result.Report.RequireNoLoss);
    }

    [Fact]
    public void Import_SplitsNestedSectionsAndRewritesCrossChapterAnchors() {
        var source = HtmlConversionDocument.Parse("<html lang='pl'><head><title>Book</title></head><body><p>Preface</p><main><section class='prose'><h1 id='one'>First</h1><p><a href='#two'>Next</a> &amp; prose</p><h2>Details</h2><ol><li>List</li></ol><h1 id='two'>Second</h1><table><tr><th scope='col'>Name</th></tr><tr><td>Value</td></tr></table><p><a href='#one'>Back</a></p></section></main></body></html>");
        EpubManuscriptResult result = EpubManuscript.ImportHtml(source, new EpubManuscriptOptions { Creator = "Author" });
        result.Report.RequireNoLoss();
        var book = result.Publication;
        Assert.Equal("pl", book.Language);
        Assert.Equal("Author", book.Creator);
        Assert.Equal(3, book.Spine.Count);
        Assert.Equal("chapter-0003.xhtml#two", (string?)book.GetContentXml("chapter-2").Descendants(Html + "a").Single().Attribute("href"));
        Assert.Equal("chapter-0002.xhtml#one", (string?)book.GetContentXml("chapter-3").Descendants(Html + "a").Single().Attribute("href"));
        Assert.Equal("prose", (string?)book.GetContentXml("chapter-3").Descendants(Html + "section").Single().Attribute("class"));
        EpubDocument reopened = book.Read(new EpubReadOptions { IncludeRawHtml = true });
        Assert.Equal(new[] { "Book", "First", "Second" }, reopened.Chapters.Select(chapter => chapter.Title));
        Assert.Contains("Value", reopened.Chapters[2].Text);
        Assert.Single(reopened.TableOfContents[1].Children);
    }

    [Fact]
    public void Import_EmptyHeadingsDoNotSplitUntilTextOrMediaIsPresent() {
        string html = "<title>Chapters</title>" + string.Concat(Enumerable.Repeat("<h1></h1>", 1024)) +
            "<p>Text</p><h1>Media</h1><img alt='Dot' src='" + ImageData + "'><h1>Final</h1><p>More</p>";
        EpubManuscriptResult result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse(html));

        result.Report.RequireNoLoss();
        Assert.Equal(3, result.Publication.Spine.Count);
        Assert.Equal(1024, result.Publication.GetContentXml("chapter-1").Descendants(Html + "h1").Count());
        Assert.Single(result.Publication.GetContentXml("chapter-2").Descendants(Html + "img"));
        Assert.Contains("More", result.Publication.GetContentXml("chapter-3").Descendants(Html + "p").Single().Value);
    }

    [Fact]
    public void Import_ParsesHtmlVoidsAndEntitiesIntoXmlWithoutLosingText() {
        var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Text</title><h1>Heading</h1><p>Zażółć&nbsp;&amp; <strong>bold</strong><br>next <ruby>漢<rt>kan</rt></ruby></p>"));
        result.Report.RequireNoLoss();
        XDocument xml = result.Publication.GetContentXml("chapter-1");
        Assert.Equal("bold", xml.Descendants(Html + "strong").Single().Value);
        Assert.Single(xml.Descendants(Html + "br"));
        Assert.Single(xml.Descendants(Html + "ruby"));
        Assert.Contains("Zażółć", result.Publication.Read().Chapters[0].Text);
    }

    [Fact]
    public void Import_ReportsActiveContentAndMissingAnchorsBeforePublishing() {
        var result = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Review</title><h1>Chapter</h1><script>danger()</script><iframe src='https://example.com'></iframe><p onclick='danger()'><a href='#absent'>Broken</a></p>"));
        Assert.False(result.Report.Succeeded);
        Assert.True(result.Report.HasLoss);
        Assert.Throws<InvalidOperationException>(result.Report.RequireNoLoss);
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "EPUB_IMPORT_ACTIVE_CONTENT_OMITTED");
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "EPUB_IMPORT_ACTIVE_ATTRIBUTE_OMITTED");
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "EPUB_IMPORT_ANCHOR_MISSING");
        Assert.DoesNotContain("danger", result.Publication.GetContentXml("chapter-1").ToString());
        result.Publication.Write().Report.RequireNoLoss();
    }

    [Fact]
    public void Import_RespectsCancellationAndRequiresReadableContent() {
        var source = HtmlConversionDocument.Parse("<title>Empty</title><body></body>");
        Assert.Throws<InvalidDataException>(() => EpubManuscript.ImportHtml(source));
        Assert.Throws<OperationCanceledException>(() => EpubManuscript.ImportHtml(source, cancellationToken: new CancellationToken(true)));
        Assert.Throws<ArgumentOutOfRangeException>(() => EpubManuscript.ImportHtml(source, new EpubManuscriptOptions { ChapterHeadingLevel = 7 }));
    }
}
