using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void AnimationLonghandOverridesEarlierShorthandName() {
        string html = KeyframeHtml("animation:1s pulse;animation-name:none");

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void LaterAnimationShorthandOverridesLonghandName() {
        string html = KeyframeHtml("animation-name:none;animation:1s pulse");

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(result.Before.Evidence);
        Assert.True(result.WasChanged);
    }

    [Fact]
    public void NegativeAnimationDurationDoesNotActivateKeyframes() {
        string html = KeyframeHtml("animation:-1s pulse");

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void NegativeAnimationDelayStillActivatesKeyframes() {
        string html = KeyframeHtml("animation:1s -1s pulse");

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(result.Before.Evidence);
        Assert.True(result.WasChanged);
    }

    [Fact]
    public void CssWideKeywordCannotBeMixedIntoAnAnimationList() {
        string html = KeyframeHtml("animation:inherit,1s pulse");

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void StatefulNegationRetainsItsStaticExclusions() {
        string dataUri = ProvenanceDataUri();
        string html = "<style>.x:not(.x,:hover){background-image:url('" + dataUri +
            "')}</style><div class='x'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Theory]
    [InlineData("image")]
    [InlineData("feImage")]
    [InlineData("use")]
    public void SvgHrefOverridesLegacyXlinkHref(string elementName) {
        string html = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'><" +
            elementName + " href='clean.png' xlink:href='" + ProvenanceDataUri() + "'/></svg>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void SvgXlinkHrefRemainsTheLegacyFallback() {
        string html = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>" +
            "<image xlink:href='" + ProvenanceDataUri() + "'/></svg>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(result.Before.Evidence);
        Assert.True(result.WasChanged);
    }

    [Fact]
    public void ResourceDiscoveryAlsoPrefersSvgHref() {
        const string clean = "https://example.test/clean.png";
        const string legacy = "https://example.test/legacy.png";
        string html = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>" +
            "<image href='" + clean + "' xlink:href='" + legacy + "'/></svg>";

        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(
            html,
            new HtmlResourcePipelineOptions { UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile() });

        Assert.Contains(manifest.Resources, resource => resource.Source == clean);
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == legacy);
    }

    [Fact]
    public void SvgHrefNormalizationDoesNotRewriteScriptText() {
        const string script = "const sample=\"<svg><image href='clean.png' xlink:href='legacy.png'/></svg>\";";
        string html = "<script>" + script + "</script>" +
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>" +
            "<image xlink:href='" + ProvenanceDataUri() + "'/></svg>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string rewritten = Encoding.UTF8.GetString(result.ToArray());

        Assert.True(result.WasChanged);
        Assert.Contains(script, rewritten, StringComparison.Ordinal);
    }

    private static string KeyframeHtml(string animationDeclaration) =>
        "<style>@keyframes pulse{from{background-image:url('" + ProvenanceDataUri() +
        "')}}.box{" + animationDeclaration + "}</style><div class='box'></div>";

    private static string ProvenanceDataUri() => "data:image/png;base64," +
        Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
}
