using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class WordHtmlAsyncResourceContracts {
    private static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZlFQAAAAASUVORK5CYII=");

    [Fact]
    public async Task AsyncImportAwaitsRemoteImagesBeforeDocumentProjection() {
        var handler = new ImageHandler();
        using var client = new HttpClient(handler);
        HtmlConversionDocument html = CreateRemoteImageDocument();
        HtmlToWordOptions options = HtmlToWordOptions.CreateTrustedDocumentProfile();
        options.AllowDocumentStylesheetLinks = false;
        options.HttpClient = client;

        HtmlToWordResult result = await html.ToWordDocumentResultAsync(options);

        using var document = result.RequireValue();
        Assert.Equal(1, handler.RequestCount);
        Assert.NotEmpty(document.Images);
    }

    [Fact]
    public void SynchronousImportRejectsInputsThatRequireNetworkAccess() {
        var handler = new ImageHandler();
        using var client = new HttpClient(handler);
        HtmlConversionDocument html = CreateRemoteImageDocument();
        HtmlToWordOptions options = HtmlToWordOptions.CreateTrustedDocumentProfile();
        options.AllowDocumentStylesheetLinks = false;
        options.HttpClient = client;

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => html.ToWordDocumentResult(options));

        Assert.Contains("offline-only", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, handler.RequestCount);
    }

    [Fact]
    public void SynchronousImportRejectsRelativeConfiguredStylesheetWithRemoteBaseUri() {
        HtmlConversionDocument html = HtmlConversionDocument.Parse(
            "<html><body><p>Remote stylesheet</p></body></html>",
            new HtmlConversionDocumentOptions {
                Trust = HtmlInputTrust.Trusted,
                BaseUri = new Uri("https://assets.example.test/pages/index.html")
            });
        HtmlToWordOptions options = HtmlToWordOptions.CreateTrustedDocumentProfile();
        options.StylesheetPaths.Add("../styles/site.css");

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => html.ToWordDocumentResult(options));

        Assert.Contains("offline-only", exception.Message, StringComparison.Ordinal);
    }

    private static HtmlConversionDocument CreateRemoteImageDocument() => HtmlConversionDocument.Parse(
        "<html><body><img src=\"https://assets.example.test/pixel.png\" alt=\"Pixel\"></body></html>",
        new HtmlConversionDocumentOptions { Trust = HtmlInputTrust.Trusted });

    private sealed class ImageHandler : HttpMessageHandler {
        internal int RequestCount { get; private set; }

        protected override async Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken) {
            RequestCount++;
            await Task.Yield();
            cancellationToken.ThrowIfCancellationRequested();
            var response = new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new ByteArrayContent(OnePixelPng)
            };
            response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("image/png");
            return response;
        }
    }
}
