using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void ExternalStylesheetManifestUsesCanonicalCssResourceDiscovery() {
        HtmlUrlPolicy policy = HtmlUrlPolicy.CreateOfficeIMOProfile();
        policy.DisallowFileUrls = false;
        policy.AllowedUrlSchemes.Add(Uri.UriSchemeFile);
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildStylesheetManifest(
            "@import 'theme.css'; @font-face { font-family: Demo; src: url('demo.woff2'); } body { background-image: url('../images/paper.png'); }",
            new Uri("file:///documents/styles/main.css"),
            new HtmlResourcePipelineOptions { ResourceUrlPolicy = policy });

        Assert.Equal(3, manifest.AllowedCount);
        Assert.Contains(manifest.Resources, resource =>
            resource.Kind == HtmlResourceKind.Stylesheet && resource.ResolvedSource.EndsWith("/documents/styles/theme.css", StringComparison.Ordinal));
        Assert.Contains(manifest.Resources, resource =>
            resource.Kind == HtmlResourceKind.Font && resource.ResolvedSource.EndsWith("/documents/styles/demo.woff2", StringComparison.Ordinal));
        Assert.Contains(manifest.Resources, resource =>
            resource.Kind == HtmlResourceKind.Image && resource.ResolvedSource.EndsWith("/documents/images/paper.png", StringComparison.Ordinal));
    }

    [Fact]
    public void ExternalStylesheetManifestEnforcesCssByteLimit() {
        Assert.Throws<HtmlDomLimitException>(() => HtmlResourcePipeline.BuildStylesheetManifest(
            "body { color: red; }",
            new Uri("https://example.test/main.css"),
            new HtmlResourcePipelineOptions {
                Limits = new HtmlConversionLimits { MaxCssBytes = 4 }
            }));
    }
}
