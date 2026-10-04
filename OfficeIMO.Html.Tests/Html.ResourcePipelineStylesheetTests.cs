using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void ManifestSelectsOnlyAlternatesInThePreferredStylesheetSet() {
        string links = string.Concat(Enumerable.Range(0, 64).Select(index =>
            $"<link rel='stylesheet' href='https://example.test/plain-{index}.css'>"))
            + "<link rel='stylesheet' title='default' href='https://example.test/default.css'>"
            + "<link rel='alternate stylesheet' title='default' href='https://example.test/selected.css'>"
            + "<link rel='alternate stylesheet' title='other' href='https://example.test/unselected.css'>";

        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("<head>" + links + "</head>");

        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/selected.css"
            && resource.Kind == HtmlResourceKind.Stylesheet);
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/unselected.css");
    }

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
    public void ExternalStylesheetManifestTreatsLegacyCdoAndCdcTokensAsImportTrivia() {
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildStylesheetManifest(
            "<!-- --> @import 'theme.css';",
            new Uri("https://example.test/styles/main.css"));

        HtmlResourceReference imported = Assert.Single(manifest.Resources);
        Assert.Equal(HtmlResourceKind.Stylesheet, imported.Kind);
        Assert.Equal("https://example.test/styles/theme.css", imported.ResolvedSource);
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

    [Fact]
    public void StylesheetDecoderHonorsDeclaredLegacyCharset() {
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(
            "@charset \"windows-1252\"; body { background-image: url('caf?.png'); }");
        bytes[Array.IndexOf(bytes, (byte)'?')] = 0xE9;

        bool decoded = HtmlResourcePipeline.TryDecodeStylesheet(bytes, contentType: null, out string css);

        Assert.True(decoded);
        Assert.Contains("café.png", css, StringComparison.Ordinal);
    }
}
