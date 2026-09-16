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
        string[] nested = discovery.DiscoverStylesheets([source]);

        Assert.Contains("https://example.test/css/nested/theme.css", nested);
        Assert.Contains("https://example.test/images/card.png", nested);
        Assert.Empty(discovery.DiscoverStylesheets([source]));
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

        Assert.Contains("https://example.test/css/nested.css", discovery.DiscoverStylesheets([stylesheet]));
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
        Assert.Contains("https://example.test/images/seal.png", discovery.DiscoverStylesheets([stylesheet]));
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
}
