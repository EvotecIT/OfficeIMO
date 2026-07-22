using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class MhtmlDocumentTests {
    [Fact]
    public async Task LoadSelectsDeclaredRootAndResolvesCidResource() {
        const string archive = "MIME-Version: 1.0\r\n" +
            "Subject: Saved page\r\n" +
            "Content-Type: multipart/related; boundary=archive; type=\"text/html\"; start=\"<root>\"\r\n\r\n" +
            "--archive\r\nContent-Type: image/png\r\nContent-ID: <logo>\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nAQID\r\n" +
            "--archive\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root>\r\n" +
            "Content-Location: https://example.test/page/index.html\r\n\r\n" +
            "<html><body><img src=\"cid:logo\"></body></html>\r\n" +
            "--archive--\r\n";
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes(archive));

        MhtmlDocument document = MhtmlDocument.Load(stream);
        HtmlResolvedResource? resolved = await document.CreateResourceResolver()(
            new HtmlRenderResourceRequest(new Uri("cid:logo"), "cid:logo", HtmlResourceKind.Image),
            CancellationToken.None);

        Assert.Contains("cid:logo", document.Html, StringComparison.Ordinal);
        Assert.Equal("root", document.RootContentId);
        Assert.Equal("https://example.test/page/index.html", document.ContentLocation);
        Assert.Equal(new Uri("https://example.test/page/index.html"), document.BaseUri);
        MhtmlResource resource = Assert.Single(document.Resources);
        Assert.Equal("logo", resource.ContentId);
        Assert.NotNull(resolved);
        Assert.Equal(new byte[] { 1, 2, 3 }, resolved!.Bytes);
        Assert.Equal("image/png", resolved.ContentType);
    }

    [Fact]
    public async Task ConfigureRenderOptionsAllowsPackageResourcesWithoutRelaxingHyperlinksOrFallbacks() {
        int fallbackCalls = 0;
        var document = new MhtmlDocument(
            "<a href='cid:logo'>link</a><img src='cid:logo'>",
            new[] { new MhtmlResource(new byte[] { 1, 2, 3 }, "image/png", contentId: "logo") },
            "file:///snapshot/page.html");
        var options = new HtmlRenderOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            ResourceResolver = (request, cancellationToken) => {
                fallbackCalls++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(new byte[] { 4, 5, 6 }, "image/png"));
            }
        };

        document.ConfigureRenderOptions(options);

        Assert.DoesNotContain("cid", options.UrlPolicy.AllowedUrlSchemes);
        Assert.DoesNotContain(Uri.UriSchemeFile, options.UrlPolicy.AllowedUrlSchemes);
        Assert.True(options.UrlPolicy.DisallowFileUrls);
        Assert.NotNull(options.ResourceUrlPolicy);
        Assert.Contains("cid", options.ResourceUrlPolicy!.AllowedUrlSchemes);
        Assert.Contains(Uri.UriSchemeFile, options.ResourceUrlPolicy.AllowedUrlSchemes);
        Assert.False(options.ResourceUrlPolicy.DisallowFileUrls);
        Assert.NotNull(options.ResourceResolver);
        HtmlResolvedResource? embedded = await options.ResourceResolver!(
            new HtmlRenderResourceRequest(new Uri("cid:logo"), "cid:logo", HtmlResourceKind.Image),
            CancellationToken.None);
        HtmlResolvedResource? missingFile = await options.ResourceResolver(
            new HtmlRenderResourceRequest(new Uri("file:///outside/secret.png"), "file:///outside/secret.png", HtmlResourceKind.Image),
            CancellationToken.None);
        Assert.NotNull(embedded);
        Assert.Null(missingFile);
        Assert.Equal(0, fallbackCalls);
    }

    [Fact]
    public void ConversionDocumentPreservesOnlyArchiveBackedCidAndFileResources() {
        var document = new MhtmlDocument(
            "<a href='cid:logo'>package link</a><img src='cid:logo'><img src='images/chart.png'><img src='file:///outside/secret.png'>",
            new[] {
                new MhtmlResource(new byte[] { 1 }, "image/png", contentId: "logo", fileName: "logo.png"),
                new MhtmlResource(new byte[] { 2 }, "image/png", contentLocation: "images/chart.png", fileName: "chart.png")
            },
            "file:///snapshot/page.html");

        string html = document.HtmlDocument.HtmlForConversion;

        Assert.Contains("src=\"cid:logo\"", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("file:///snapshot/images/chart.png", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("file:///outside/secret.png", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("href=\"cid:logo\"", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ConstructedArchiveRoundTripsUnreferencedRelatedResource() {
        var resource = new MhtmlResource(Encoding.UTF8.GetBytes("body { color: black; }"),
            "text/css", contentLocation: "styles/site.css", fileName: "site.css");
        var document = new MhtmlDocument("<html><body>saved</body></html>", new[] { resource },
            "https://example.test/page/index.html", "root", "Saved page");

        byte[] bytes = document.ToBytes();
        string serialized = Encoding.ASCII.GetString(bytes);
        using var stream = new MemoryStream(bytes);
        MhtmlDocument roundTrip = MhtmlDocument.Load(stream);

        Assert.Contains("multipart/related", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("type=\"text/html\"", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("start=\"<root>\"", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Saved page", roundTrip.Subject);
        Assert.Equal("root", roundTrip.RootContentId);
        Assert.Equal("styles/site.css", Assert.Single(roundTrip.Resources).ContentLocation);
    }

    [Fact]
    public void ConstructedArchiveWithoutResourcesStillUsesRelatedContainer() {
        var document = new MhtmlDocument("<html><body>standalone</body></html>", rootContentId: "root");

        byte[] bytes = document.ToBytes();
        using var stream = new MemoryStream(bytes);
        MhtmlDocument roundTrip = MhtmlDocument.Load(stream);

        Assert.Equal("root", roundTrip.RootContentId);
        Assert.Empty(roundTrip.Resources);
    }

    [Fact]
    public void LoadRejectsOrdinaryEmailMessage() {
        const string message = "Subject: ordinary\r\nContent-Type: text/html; charset=utf-8\r\n\r\n<p>mail</p>\r\n";
        using var stream = new MemoryStream(Encoding.ASCII.GetBytes(message));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => MhtmlDocument.Load(stream));

        Assert.Contains("multipart/related", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
