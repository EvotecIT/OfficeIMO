using System.Collections.Concurrent;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private const string GeneratedResourcePixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlGeneratedBackgroundResources_ImportedCascadePreservesBothOriginsAndOneRewrite(bool asynchronous) {
        const string html = "<base href='https://document.example.test/base/'><style>:root{--document-image:url(document.png)}</style>"
            + "<link rel='stylesheet' href='https://assets.example.test/css/main.css'><p>Body</p>";
        const string main = "@import 'nested/page.css';";
        const string imported = "@page{--sheet-image:url('../images/paper.png');@top-center{content:'';width:24px;height:12px;"
            + "background-image:var(--document-image),var(--sheet-image);background-repeat:no-repeat}}";
        var requests = new ConcurrentQueue<string>();
        HtmlRenderOptions options = CreateGeneratedResourceOptions(request => {
            requests.Enqueue(request.Uri.AbsoluteUri);
            return request.Uri.AbsolutePath.EndsWith("main.css", StringComparison.Ordinal)
                ? new HtmlResolvedResource(Encoding.UTF8.GetBytes(main), "text/css")
                : request.Uri.AbsolutePath.EndsWith("page.css", StringComparison.Ordinal)
                    ? new HtmlResolvedResource(Encoding.UTF8.GetBytes(imported), "text/css")
                    : new HtmlResolvedResource(Convert.FromBase64String(GeneratedResourcePixel), "image/png");
        });
        options.ResourceUrlPolicy!.ResolvedUrlTransform = source => source + (source.Contains('?') ? "&" : "?") + "signed=1";

        HtmlRenderDocument rendered = asynchronous
            ? await HtmlRenderEngine.RenderAsync(html, options) : HtmlRenderEngine.Render(html, options);

        Assert.Equal(new[] {
            "https://assets.example.test/css/images/paper.png?signed=1",
            "https://assets.example.test/css/main.css?signed=1",
            "https://assets.example.test/css/nested/page.css?signed=1",
            "https://document.example.test/base/document.png?signed=1"
        }, requests.OrderBy(source => source, StringComparer.Ordinal));
        Assert.Equal(2, rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderImage>().Count());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlGeneratedBackgroundResources_DoesNotRequestSuppressedExternalPageSources(bool asynchronous) {
        const string css = "@page{--image:url(unused-variable.png);"
            + "@top-left{content:none;background-image:url(none.png)}"
            + "@top-center{content:normal;background-image:var(--image)}"
            + "@bottom-center{content:'';background-image:url(overridden.png);background-image:none}"
            + "@left-middle{content:'';background-image:url(vertical.png)}}";
        var requests = new ConcurrentQueue<string>();
        HtmlRenderOptions options = CreateGeneratedResourceOptions(request => {
            requests.Enqueue(request.Uri.AbsoluteUri);
            return new HtmlResolvedResource(Encoding.UTF8.GetBytes(css), "text/css");
        });
        const string html = "<link rel='stylesheet' href='https://example.test/main.css'><p>Body</p>";

        if (asynchronous) await HtmlRenderEngine.RenderAsync(html, options);
        else HtmlRenderEngine.Render(html, options);

        Assert.Equal("https://example.test/main.css", Assert.Single(requests));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task HtmlGeneratedBackgroundResources_RejectsOnlyEffectiveExternalPagePolicyUse(bool asynchronous, bool generated) {
        string css = "@page{--image:url(file:///private/blocked.png);@top-center{content:"
            + (generated ? "''" : "none") + ";width:20px;height:12px;background-image:var(--image)}}";
        var requests = new ConcurrentQueue<string>();
        HtmlRenderOptions options = CreateGeneratedResourceOptions(request => {
            requests.Enqueue(request.Uri.AbsoluteUri);
            return new HtmlResolvedResource(Encoding.UTF8.GetBytes(css), "text/css");
        });
        const string html = "<link rel='stylesheet' href='https://example.test/main.css'><p>Body</p>";

        HtmlRenderDocument rendered = asynchronous
            ? await HtmlRenderEngine.RenderAsync(html, options) : HtmlRenderEngine.Render(html, options);

        Assert.Equal("https://example.test/main.css", Assert.Single(requests));
        Assert.Equal(generated, rendered.Diagnostics.Any(diagnostic => diagnostic.Code == "ImageResourceRejectedByPolicy"
            && diagnostic.Source == "file:///private/blocked.png"));
    }

    [Theory]
    [InlineData(false, 3)]
    [InlineData(false, 4)]
    [InlineData(true, 3)]
    [InlineData(true, 4)]
    public async Task HtmlGeneratedBackgroundResources_SharesRequestBudgetAndPreviouslyLoadedImage(bool asynchronous, int requestLimit) {
        const string css = "@page{@top-center{content:'';width:20px;height:12px;"
            + "background-image:url(shared.png),url(late.png),url(late.png);background-repeat:no-repeat}}";
        var requests = new ConcurrentQueue<string>();
        HtmlRenderOptions options = CreateGeneratedResourceOptions(request => {
            requests.Enqueue(request.Uri.AbsoluteUri);
            if (request.Uri.AbsolutePath.EndsWith("missing.png", StringComparison.Ordinal)) return null;
            return request.Kind == HtmlResourceKind.Stylesheet
                ? new HtmlResolvedResource(Encoding.UTF8.GetBytes(css), "text/css")
                : new HtmlResolvedResource(Convert.FromBase64String(GeneratedResourcePixel), "image/png");
        });
        options.MaxResourceRequests = requestLimit;
        options.MaxResourceCount = requestLimit;
        const string html = "<link rel='stylesheet' href='https://example.test/main.css'>"
            + "<img src='https://example.test/shared.png' width='1' height='1' alt='pixel'>"
            + "<img src='https://example.test/missing.png' width='1' height='1' alt='unavailable'>";

        HtmlRenderDocument rendered = asynchronous
            ? await HtmlRenderEngine.RenderAsync(html, options) : HtmlRenderEngine.Render(html, options);

        Assert.Equal(requestLimit, requests.Count);
        Assert.Single(requests, source => source == "https://example.test/shared.png");
        Assert.Equal(requestLimit == 4, requests.Contains("https://example.test/late.png"));
        Assert.Equal(requestLimit == 3, rendered.Diagnostics.Any(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceRequestLimitExceeded));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlGeneratedBackgroundResources_MissingSignedExternalImageRemainsVisibleLoss(bool asynchronous) {
        const string css = "@page{--image:url(missing.png);@top-center{content:'';width:20px;height:12px;background-image:var(--image)}}";
        var requests = new ConcurrentQueue<string>();
        HtmlRenderOptions options = CreateGeneratedResourceOptions(request => {
            requests.Enqueue(request.Uri.AbsoluteUri);
            return request.Kind == HtmlResourceKind.Stylesheet ? new HtmlResolvedResource(Encoding.UTF8.GetBytes(css), "text/css") : null;
        });
        options.ResourceUrlPolicy!.ResolvedUrlTransform = source => source + (source.Contains('?') ? "&" : "?") + "signed=1";
        const string html = "<link rel='stylesheet' href='https://example.test/main.css'><p>Body</p>";

        HtmlRenderDocument rendered = asynchronous
            ? await HtmlRenderEngine.RenderAsync(html, options) : HtmlRenderEngine.Render(html, options);

        Assert.Equal(new[] { "https://example.test/main.css?signed=1", "https://example.test/missing.png?signed=1" }, requests);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable
            && diagnostic.Source == "missing.png" && diagnostic.Severity == HtmlDiagnosticSeverity.Warning
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlGeneratedBackgroundResources_EmbeddedViewportCannotInvokeParentResolver() {
        int calls = 0;
        HtmlRenderOptions parentOptions = CreateGeneratedResourceOptions(_ => { calls++; return null; });
        HtmlResourceSession parentResources = HtmlResourceSession.Resolve(new HtmlResourceManifest(), parentOptions);
        HtmlConversionDocument nested = HtmlConversionDocument.Parse("<style>@page{@top-center{content:'';width:20px;height:12px;"
            + "background-image:url(https://example.test/not-prefetched.png)}}</style><p>Nested</p>");

        HtmlRenderEngine.RenderEmbedded(nested, new HtmlRenderOptions { AllowSystemFontFallback = false },
            parentResources, CancellationToken.None);

        Assert.Equal(0, calls);
        Assert.Equal(0, parentResources.ResolverRequestCount);
    }

    private static HtmlRenderOptions CreateGeneratedResourceOptions(Func<HtmlRenderResourceRequest, HtmlResolvedResource?> resolver) => new() {
        AllowSystemFontFallback = false,
        ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile(),
        ResourceResolver = (request, _) => Task.FromResult(resolver(request)),
        SynchronousResourceResolver = (HtmlRenderResourceRequest request, CancellationToken _, out HtmlResolvedResource? resource) => {
            resource = resolver(request);
            return true;
        }
    };
}
