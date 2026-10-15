using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Theory]
    [InlineData("image-set(cross-fade(image-set(\"x.png\" 1x),url(y.png),50%) 1x)", false)]
    [InlineData("image-set(cross-fade(-webkit-image-set(\"x.png\" 1x),url(y.png),50%) 1x)", false)]
    [InlineData("image-set(cross-fade(\\69mage-set(\"x.png\" 1x),url(y.png),50%) 1x)", false)]
    [InlineData("image-set(\"image-set(x.png)\" 1x)", true)]
    [InlineData("image-set(url(\"https://example.test/image-set(x.png).png\") 1x)", true)]
    [InlineData("image-set(url(https://example.test/image-set\\(x.png\\).png) 1x)", true)]
    public void ImageSetSyntaxRejectsNestedFunctionsWithoutInspectingUrlOrStringPayloads(string value, bool valid) {
        Assert.Equal(valid, HtmlResourcePipeline.IsValidImageSetSyntax(value, _ => true));
    }

    [Theory]
    [InlineData("image-set(\"\")", false)]
    [InlineData("image-set(\"https://example.test/image.png\" type(\"\"))", false)]
    [InlineData("image-set(\"https://example.test/image.png\" type(\"image/png\"))", true)]
    public void ImageSetSyntaxAdmissionDoesNotExpandResourceExtractionEligibility(string value, bool discoversImage) {
        Assert.True(HtmlResourcePipeline.IsValidImageSetSyntax(value, _ => true));
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(
            "<style>.box{background-image:" + value + "}</style><div class='box'></div>");
        Assert.Equal(discoversImage, manifest.Resources.Any(resource => resource.Kind == HtmlResourceKind.Image));
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
