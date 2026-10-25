using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void GeneratedBackgroundManifestUsesPageInheritanceAndNamedCascades() {
        const string css = ":root{--image:url(root.png)}"
            + "@page first{--image:url(first.png);@top-center{content:'';background-image:var(--image)}}"
            + "@page second{--image:url(second.png);@top-center{content:'';background-image:var(--image)}}";

        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("<style>" + css + "</style><p>Body</p>",
            new HtmlResourcePipelineOptions { BaseUri = new Uri("https://example.test/document/") });

        Assert.Equal(new[] { "https://example.test/document/first.png", "https://example.test/document/second.png" },
            manifest.Resources.Where(resource => resource.Kind == HtmlResourceKind.Image)
                .Select(resource => resource.ResolvedSource).OrderBy(source => source, StringComparer.Ordinal));
    }

    [Theory]
    [InlineData("none", 0)]
    [InlineData("normal", 0)]
    [InlineData("''", 1)]
    public void GeneratedBackgroundManifestRequiresGeneratedContent(string content, int imageCount) {
        string css = "@page{--image:url(inherited.png);@top-center{content:" + content
            + ";background-image:var(--image)}}";

        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("<style>" + css + "</style>");

        Assert.Equal(imageCount, manifest.Resources.Count(resource => resource.Kind == HtmlResourceKind.Image));
    }

    [Fact]
    public void GeneratedBackgroundManifestKeepsEffectiveSideImagesAndSuppressesOverriddenSources() {
        const string css = "@page{@top-center{content:'';background-image:url(overridden.png);background-image:none}"
            + "@left-middle{content:'';background-image:url(vertical.png)}}"
            + ".used{background-image:url(ordinary.png)}";

        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("<style>" + css + "</style><div class='used'></div>");

        Assert.Equal(new[] { "ordinary.png", "vertical.png" }, manifest.Resources
            .Where(resource => resource.Kind == HtmlResourceKind.Image).Select(resource => resource.Source)
            .OrderBy(source => source, StringComparer.Ordinal));
    }

    [Fact]
    public void GeneratedBackgroundManifestPreservesEffectiveBaseAndResourcePolicy() {
        const string html = "<base href='https://example.test/document/'><style>"
            + ":root{--image:url(root.png)}@page{@top-center{content:'';background-image:var(--image)}"
            + "@bottom-center{content:'';background-image:url(file:///private/blocked.png)}}</style>";

        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(html, new HtmlResourcePipelineOptions {
            BaseUri = new Uri("https://fallback.example.test/"),
            ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile()
        });

        Assert.Contains(manifest.Resources, resource => resource.IsAllowed && resource.ResolvedSource == "https://example.test/document/root.png");
        Assert.Contains(manifest.Resources, resource => !resource.IsAllowed && resource.Source == "file:///private/blocked.png"
            && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
    }
}
