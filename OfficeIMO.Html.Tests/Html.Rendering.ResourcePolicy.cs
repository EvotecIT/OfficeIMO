using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlResourceSession_RejectsPermissiveManifestBeforeCallingResolver(bool asynchronous) {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/site.css'>"
            + "<style>@font-face{font-family:Remote;src:url(https://assets.example.test/font.ttf)}</style>"
            + "<img src='https://assets.example.test/chart.png'>");
        Assert.Equal(3, source.ResourceManifest.AllowedCount);
        int requests = 0;
        var options = new HtmlRenderOptions {
            ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile(),
            ResourceResolver = (request, token) => {
                requests++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Array.Empty<byte>(), "text/css"));
            }
        };
        options.SynchronousResourceResolver = (HtmlRenderResourceRequest request, CancellationToken token, out HtmlResolvedResource? resource) => {
            requests++;
            resource = new HtmlResolvedResource(Array.Empty<byte>(), "text/css", request.Uri);
            return true;
        };

        HtmlResourceSession session = asynchronous
            ? await HtmlResourceSession.ResolveAsync(source.ResourceManifest, options)
            : HtmlResourceSession.Resolve(source.ResourceManifest, options);

        Assert.Equal(0, requests);
        Assert.Equal(0, session.ResolverRequestCount);
        Assert.Empty(session.Resources);
        Assert.Equal(new[] { "FontResourceRejectedByPolicy", "ImageResourceRejectedByPolicy", "StylesheetResourceRejectedByPolicy" },
            session.Diagnostics.Select(diagnostic => diagnostic.Code).OrderBy(code => code).ToArray());
        Assert.All(session.Diagnostics, diagnostic => Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind));
    }

    [Fact]
    public async Task HtmlResourceSession_RechecksHostPolicyBeforeInitialAndImportedRequests() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://allowed.example.test/site.css'>"
            + "<img src='https://blocked.example.test/image.png'>");
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();
        policy.ResolvedUrlTransform = url => new Uri(url).Host == "allowed.example.test" ? url : null;
        var requests = new List<Uri>();
        var options = new HtmlRenderOptions {
            ResourceUrlPolicy = policy,
            ResourceResolver = (request, token) => {
                requests.Add(request.Uri);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                    Encoding.UTF8.GetBytes("@import url('https://blocked.example.test/import.css');"), "text/css"));
            }
        };

        HtmlResourceSession session = await HtmlResourceSession.ResolveAsync(source.ResourceManifest, options);

        Assert.Equal("https://allowed.example.test/site.css", Assert.Single(requests).AbsoluteUri);
        Assert.Single(session.Resources);
        Assert.Contains(session.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceRejectedByPolicy");
        Assert.Contains(session.Diagnostics, diagnostic => diagnostic.Code == "StylesheetResourceRejectedByPolicy");
    }

    [Fact]
    public async Task HtmlResourceSession_WaitsForUncooperativeCallbackAndRejectsLateContent() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/site.css'>");
        var cancelled = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var released = new TaskCompletionSource<HtmlResolvedResource?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var options = new HtmlRenderOptions {
            ResourceTimeout = TimeSpan.FromMilliseconds(30),
            ResourceResolver = async (request, token) => {
                using (token.Register(() => cancelled.TrySetResult(true))) {
                    return await released.Task;
                }
            }
        };
        Task<HtmlResourceSession> operation = HtmlResourceSession.ResolveAsync(source.ResourceManifest, options);
        try {
            Assert.Same(cancelled.Task, await Task.WhenAny(cancelled.Task, Task.Delay(TimeSpan.FromSeconds(10))));
            Assert.False(operation.IsCompleted);
        } finally {
            released.TrySetResult(new HtmlResolvedResource(Encoding.UTF8.GetBytes("p{color:red}"), "text/css"));
        }

        HtmlResourceSession session = await operation;

        Assert.Empty(session.Resources);
        Assert.Contains(session.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceTimeout
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }
}
