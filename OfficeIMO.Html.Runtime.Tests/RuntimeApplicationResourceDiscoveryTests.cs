using OfficeIMO.Html.Runtime.Rendering;
using Xunit;

namespace OfficeIMO.Html.Runtime.Tests;

public sealed class RuntimeApplicationResourceDiscoveryTests {
    [Fact]
    public void DocumentDiscoveryUsesOwnedHtmlCssPlannerWithoutFetchingHyperlinks() {
        const string html = """
            <html><head><base href="https://example.test/site/">
            <link rel="stylesheet" href="app.css"><script type="module" src="app.js"></script>
            <style>@import url('theme.css'); .card { background-image: url('card.png') }</style>
            </head><body><a href="private">link</a><div class="card"></div><img src="picture.png"></body></html>
            """;
        var discovery = new HtmlApplicationResourceDiscovery();

        string[] urls = discovery.DiscoverDocument(html, new Uri("https://example.test/"), []);

        Assert.Contains("https://example.test/site/app.css", urls);
        Assert.Contains("https://example.test/site/app.js", urls);
        Assert.Contains("https://example.test/site/theme.css", urls);
        Assert.Contains("https://example.test/site/card.png", urls);
        Assert.Contains("https://example.test/site/picture.png", urls);
        Assert.DoesNotContain("https://example.test/site/private", urls);
    }

    [Fact]
    public void ExternalStylesheetDiscoveryUsesFinalUrlAndRecursesWithoutDuplicateFetches() {
        var discovery = new HtmlApplicationResourceDiscovery();
        Uri document = new("https://example.test/page");
        Uri stylesheet = new("https://example.test/css/app.css");
        Assert.Contains(stylesheet.AbsoluteUri, discovery.DiscoverDocument(
            "<link rel='stylesheet' href='/css/app.css'>", document, []));

        var source = HtmlRuntimeResource.FromText(stylesheet,
            "@import url('nested/theme.css'); .card { background-image: url('../images/card.png') }", "text/css");
        string[] nested = discovery.DiscoverResources([source]);

        Assert.Contains("https://example.test/css/nested/theme.css", nested);
        Assert.Contains("https://example.test/images/card.png", nested);
        Assert.Empty(discovery.DiscoverResources([source]));
    }

    [Fact]
    public void SuppliedStylesheetIsParsedWithoutBeingRequestedAgain() {
        Uri document = new("https://example.test/page");
        Uri stylesheet = new("https://example.test/css/app.css");
        var supplied = HtmlRuntimeResource.FromText(stylesheet,
            "@import url('theme.css');", "text/css");
        var discovery = new HtmlApplicationResourceDiscovery();

        string[] urls = discovery.DiscoverDocument("<link rel='stylesheet' href='/css/app.css'>", document, [supplied]);

        Assert.DoesNotContain(stylesheet.AbsoluteUri, urls);
        Assert.Contains("https://example.test/css/theme.css", urls);
    }

    [Fact]
    public void RuntimeDiscoveredCssFollowsItsImports() {
        var discovery = new HtmlApplicationResourceDiscovery();
        Uri document = new("https://example.test/page");
        Assert.Empty(discovery.DiscoverDocument("<p>Ready</p>", document, []));
        var stylesheet = HtmlRuntimeResource.FromText(new Uri(document, "/css/dynamic.css"),
            "@import url('nested.css');", "text/css; charset=utf-8");

        Assert.Contains("https://example.test/css/nested.css", discovery.DiscoverResources([stylesheet]));
    }

    [Fact]
    public void PrintOnlyStylesheetAndImageAreDiscoveredForPdfOutput() {
        var discovery = new HtmlApplicationResourceDiscovery();
        Uri document = new("https://example.test/report");
        string[] urls = discovery.DiscoverDocument("""
            <link rel='stylesheet' media='print' href='/css/print.css'>
            <style>@media print { .invoice { background-image: url('/images/stamp.png') } }</style>
            <div class='invoice'>Report</div>
            """, document, []);

        Assert.Contains("https://example.test/css/print.css", urls);
        Assert.Contains("https://example.test/images/stamp.png", urls);

        var stylesheet = HtmlRuntimeResource.FromText(new Uri(document, "/css/print.css"),
            "@media print { .invoice { background-image: url('../images/seal.png') } }", "text/css");
        Assert.Contains("https://example.test/images/seal.png", discovery.DiscoverResources([stylesheet]));
    }

    [Fact]
    public void ResponsivePictureDiscoveryUsesOnlyTheActiveScreenAndPrintSource() {
        var discovery = new HtmlApplicationResourceDiscovery();
        string[] urls = discovery.DiscoverDocument("""
            <picture>
              <source media="(max-width: 900px)" type="image/svg+xml" srcset="/wide.svg">
              <source media="(min-width: 901px)" type="image/svg+xml" srcset="/narrow.svg">
              <img src="/fallback.svg" alt="fixture">
            </picture>
            """, new Uri("https://example.test/report"), []);

        Assert.Equal(new[] { "https://example.test/wide.svg" }, urls);
    }

    [Fact]
    public void FrameDocumentsAreDiscoveredAndTheirStaticResourcesUseTheFinalDocumentUrl() {
        var discovery = new HtmlApplicationResourceDiscovery();
        Uri page = new("https://example.test/reports/index.html");
        Uri frame = new("https://example.test/frames/detail.html");
        Uri finalFrame = new("https://assets.example.test/frame/redirected/detail.html");

        Assert.Equal(new[] { frame.AbsoluteUri }, discovery.DiscoverDocument(
            "<iframe src='../frames/detail.html'></iframe>", page, []));

        var frameDocument = new HtmlRuntimeResource(frame, System.Text.Encoding.UTF8.GetBytes("""
            <link rel="stylesheet" href="frame.css">
            <script src="frame.js"></script>
            <img src="frame.png">
            <iframe src="nested/child.html"></iframe>
            """), "text/html; charset=utf-8", finalUrl: finalFrame, redirectCount: 1);
        string[] nested = discovery.DiscoverResources([frameDocument]);

        Assert.Equal(new[] {
            "https://assets.example.test/frame/redirected/frame.css",
            "https://assets.example.test/frame/redirected/frame.js",
            "https://assets.example.test/frame/redirected/frame.png",
            "https://assets.example.test/frame/redirected/nested/child.html"
        }, nested);
        Assert.Empty(discovery.DiscoverResources([frameDocument]));
    }

    [Fact]
    public void SuppliedNestedFramesAreProcessedToAFixpointRegardlessOfResourceOrder() {
        var discovery = new HtmlApplicationResourceDiscovery();
        Uri page = new("https://example.test/index.html");
        Uri child = new("https://example.test/frames/child.html");
        Uri grandchild = new("https://example.test/frames/nested/grandchild.html");
        var grandchildDocument = HtmlRuntimeResource.FromText(grandchild,
            "<img src='grandchild.png'>", "text/html; charset=utf-8");
        var childDocument = HtmlRuntimeResource.FromText(child,
            "<iframe src='nested/grandchild.html'></iframe>", "text/html; charset=utf-8");

        string[] discovered = discovery.DiscoverDocument("<iframe src='frames/child.html'></iframe>", page,
            [grandchildDocument, childDocument]);

        Assert.Equal(new[] { "https://example.test/frames/nested/grandchild.png" }, discovered);
    }

    [Fact]
    public void DiscoveredFramesRejectInvalidResponsesAndEncodings() {
        Uri page = new("https://example.test/index.html");
        Uri frame = new("https://example.test/frame.html");
        HtmlRuntimeResource[] invalid = [
            new(frame, System.Text.Encoding.UTF8.GetBytes("<p>wrong charset</p>"), "text/html; charset=windows-1252"),
            new(frame, [0xC3, 0x28], "text/html; charset=utf-8"),
            new(frame, [], "text/html; charset=utf-8"),
            new(frame, System.Text.Encoding.UTF8.GetBytes("<p>not found</p>"), "text/html; charset=utf-8", statusCode: 404),
            new(frame, System.Text.Encoding.UTF8.GetBytes("<p>xhtml</p>"), "application/xhtml+xml; charset=utf-8")
        ];

        foreach (HtmlRuntimeResource resource in invalid) {
            var discovery = new HtmlApplicationResourceDiscovery();
            Assert.Throws<HtmlScriptRuntimeException>(() =>
                discovery.DiscoverDocument("<iframe src='frame.html'></iframe>", page, [resource]));
        }
    }
}
