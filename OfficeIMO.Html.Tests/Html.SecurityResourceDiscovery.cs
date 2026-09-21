using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void ResourceDiscoveryScansManyCssUrlsAndIgnoresQuotedText() {
        string rules = string.Concat(Enumerable.Repeat(
            ".box{background-image:url('https://example.test/image.png')}", 2048));
        string css = ".box{content:\"escaped \\\" url(https://example.test/ignored.png)\"}" + rules;

        Assert.False(HtmlResourcePipeline.HasStylesheetUrlResources(css));
    }

    [Fact]
    public void ProvenanceDiscoveryScansManyUrlsInOneDeclaration() {
        string css = ".box{background-image:" + string.Join(",",
            Enumerable.Repeat("url('data:image/png;base64,AAAA')", 2048)) + "}";

        Assert.Equal(2048, HtmlResourcePipeline.EnumerateProvenanceCssImageReferences(css).Count());
    }

    [Theory]
    [InlineData("animation:1s \\66 oo", true)]
    [InlineData("animation:1s other", false)]
    public void ResourceDiscoveryResolvesEscapedActiveKeyframeNames(string animation, bool expected) {
        const string image = "https://example.test/keyframe.png";
        string css = "@keyframes foo{from{background-image:url('" + image +
            "')}}.box{" + animation + "}";
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(
            "<style>" + css + "</style><div class='box'></div>");

        Assert.Equal(expected, manifest.Resources.Any(resource =>
            resource.Kind == HtmlResourceKind.Image && resource.Source == image));
    }
}
