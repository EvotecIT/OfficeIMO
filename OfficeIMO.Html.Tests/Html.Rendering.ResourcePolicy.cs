using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public async Task HtmlResourceSession_PreservesAliasesForRewrittenDuplicates(bool asynchronous, int maxResources) {
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(
            "<link rel='stylesheet' href='https://assets.example.test/a.css'>"
            + "<link rel='stylesheet' href='https://assets.example.test/b.css'>");
        var policy = HtmlUrlPolicy.CreateWebResourceProfile();
        policy.ResolvedUrlTransform = _ => "https://assets.example.test/shared.css";
        int requests = 0;
        var options = new HtmlRenderOptions {
            ResourceUrlPolicy = policy, MaxResourceCount = maxResources, MaxResourceRequests = maxResources,
            ResourceResolver = (request, token) => {
                requests++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Array.Empty<byte>(), "text/css"));
            }
        };
        options.SynchronousResourceResolver = (HtmlRenderResourceRequest request, CancellationToken token, out HtmlResolvedResource? resource) => {
            requests++;
            resource = new HtmlResolvedResource(Array.Empty<byte>(), "text/css");
            return true;
        };

        HtmlResourceSession session = asynchronous
            ? await HtmlResourceSession.ResolveAsync(manifest, options)
            : HtmlResourceSession.Resolve(manifest, options);

        Assert.Equal(1, requests);
        Assert.Equal(1, session.ResolverRequestCount);
        Assert.Equal(1, session.AcceptedResourceCount);
        Assert.Equal("https://assets.example.test/shared.css", Assert.Single(session.Resources).CanonicalSource);
        foreach (HtmlResourceReference reference in manifest.Resources) {
            Assert.True(session.TryGet(reference.Source, reference.ResolvedSource, out _));
        }
        Assert.Empty(session.Diagnostics);
    }

    [Fact]
    public async Task HtmlRendering_DoesNotDispatchTransformTargetRejectedByDocumentPolicy() {
        var documentPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();
        documentPolicy.ResolvedUrlTransform = value => new Uri(value).Host == "assets.example.test" ? value : null;
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/site.css'>",
            new HtmlConversionDocumentOptions { ResourceUrlPolicy = documentPolicy });
        var operationPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();
        operationPolicy.ResolvedUrlTransform = value => value.Replace("assets.example.test", "blocked.example.test");
        int requests = 0;
        HtmlRenderDocument result = await HtmlRenderEngine.RenderAsync(source, new HtmlRenderOptions {
            ResourceUrlPolicy = operationPolicy,
            ResourceResolver = (request, token) => {
                requests++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Array.Empty<byte>(), "text/css", request.Uri));
            }
        });

        Assert.Equal(0, requests);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "StylesheetResourceRejectedByPolicy");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlResourceSession_PreservesManifestRewriteUnderLaterPolicy(bool guardHost) {
        var planningPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();
        planningPolicy.ResolvedUrlTransform = value => value.Replace("untrusted.example.test", "assets.example.test");
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(
            "<link rel='stylesheet' href='https://untrusted.example.test/site.css'>",
            new HtmlResourcePipelineOptions { ResourceUrlPolicy = planningPolicy });
        var operationPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();
        if (guardHost) operationPolicy.ResolvedUrlTransform = value => new Uri(value).Host == "assets.example.test" ? value : null;
        var requested = new List<Uri>();
        await HtmlResourceSession.ResolveAsync(manifest, new HtmlRenderOptions {
            ResourceUrlPolicy = operationPolicy,
            ResourceResolver = (request, token) => {
                requested.Add(request.Uri);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Array.Empty<byte>(), "text/css", request.Uri));
            }
        });

        Assert.Equal("https://assets.example.test/site.css", Assert.Single(requested).AbsoluteUri);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlRendering_IntersectsSignedResourcePoliciesWithoutRepeatingPlanningTransform(bool reuseSigner) {
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();
        policy.ResolvedUrlTransform = value => value + "&auth=1";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/site.css?asset=1'><p>Text</p>",
            new HtmlConversionDocumentOptions { ResourceUrlPolicy = policy });
        var operationPolicy = reuseSigner ? policy.Clone() : HtmlUrlPolicy.CreateWebOnlyProfile();
        if (!reuseSigner) operationPolicy.ResolvedUrlTransform = value => new Uri(value).Host == "assets.example.test" ? value : null;
        var requested = new List<Uri>();
        await HtmlRenderEngine.RenderAsync(source, new HtmlRenderOptions {
            ResourceUrlPolicy = operationPolicy,
            ResourceResolver = (request, token) => {
                requested.Add(request.Uri);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Array.Empty<byte>(), "text/css", request.Uri));
            }
        });

        Assert.Equal("https://assets.example.test/site.css?asset=1&auth=1", Assert.Single(requested).AbsoluteUri);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task HtmlResourceSession_AppliesUrlTransformToOriginalSource(bool asynchronous, bool transformDuringPlanning) {
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();
        policy.ResolvedUrlTransform = value => value + "&auth=1";
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(
            "<base href='https://assets.example.test/files/'><link rel='stylesheet' href='site.css?asset=1'>",
            new HtmlResourcePipelineOptions {
                ResourceUrlPolicy = transformDuringPlanning ? policy : HtmlUrlPolicy.CreateWebOnlyProfile()
            });
        var requested = new List<Uri>();
        var options = new HtmlRenderOptions { ResourceUrlPolicy = policy };
        options.ResourceResolver = (request, token) => {
            requested.Add(request.Uri);
            return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(Array.Empty<byte>(), "text/css", request.Uri));
        };
        options.SynchronousResourceResolver = (HtmlRenderResourceRequest request, CancellationToken token, out HtmlResolvedResource? resource) => {
            requested.Add(request.Uri);
            resource = new HtmlResolvedResource(Array.Empty<byte>(), "text/css", request.Uri);
            return true;
        };

        HtmlResourceSession session = asynchronous
            ? await HtmlResourceSession.ResolveAsync(manifest, options)
            : HtmlResourceSession.Resolve(manifest, options);

        Assert.Equal("https://assets.example.test/files/site.css?asset=1&auth=1", Assert.Single(requested).AbsoluteUri);
        Assert.Equal(requested[0].AbsoluteUri, Assert.Single(session.Resources).CanonicalSource);
        Assert.Empty(session.Diagnostics);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlResourceSession_TransformsImportedSourceAgainstStylesheetBase(bool reportFinalUri) {
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();
        policy.ResolvedUrlTransform = value => value + "&auth=1";
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(
            "<link rel='stylesheet' href='https://assets.example.test/styles/site.css?asset=1'>");
        var requested = new List<Uri>();
        var options = new HtmlRenderOptions {
            ResourceUrlPolicy = policy,
            ResourceResolver = (request, token) => {
                requested.Add(request.Uri);
                string css = requested.Count == 1 ? "@import url('../shared/import.css?asset=2');" : string.Empty;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                    Encoding.UTF8.GetBytes(css), "text/css", reportFinalUri ? request.Uri : null));
            }
        };

        HtmlResourceSession session = await HtmlResourceSession.ResolveAsync(manifest, options);

        Assert.Equal(new[] {
            "https://assets.example.test/styles/site.css?asset=1&auth=1",
            "https://assets.example.test/shared/import.css?asset=2&auth=1"
        }, requested.Select(uri => uri.AbsoluteUri).ToArray());
        Assert.Equal(2, session.Resources.Count);
        Assert.Equal(requested.Select(uri => uri.AbsoluteUri), session.Resources.Select(resource => resource.CanonicalSource));
        Assert.Empty(session.Diagnostics);
    }

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
