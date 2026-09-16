using OfficeIMO.Html;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlResponsiveImageTests {
    [Fact]
    public void SelectorUsesDensityAndKeepsTheFirstDuplicateDensity() {
        HtmlResponsiveImageSelection selected = HtmlResponsiveImageSelector.Select(
            "one.png 1x, two.png 2x, duplicate.png 2x, four.png 4x", null, null,
            new HtmlResponsiveImageSelectionOptions { DevicePixelRatio = 2D });

        Assert.True(selected.HasValue);
        Assert.Equal("two.png", selected.Candidate.Url);
        Assert.Equal(2D, selected.EffectiveDensity);
    }

    [Fact]
    public void SelectorAddsTheDefaultSourceAsOneXForDensitySets() {
        HtmlResponsiveImageSelection selected = HtmlResponsiveImageSelector.Select(
            "large.png 2x", null, "fallback.png",
            new HtmlResponsiveImageSelectionOptions { DevicePixelRatio = 1D });

        Assert.True(selected.HasValue);
        Assert.Equal("fallback.png", selected.Candidate.Url);
        Assert.Equal(1D, selected.EffectiveDensity);
        Assert.True(selected.UsesDefaultSource);
    }

    [Fact]
    public void SelectorPreservesAuthoredCandidateOriginWhenItMatchesTheDefaultUrl() {
        HtmlResponsiveImageSelection selected = HtmlResponsiveImageSelector.Select(
            "same.png, high.png 2x", null, "same.png",
            new HtmlResponsiveImageSelectionOptions { DevicePixelRatio = 1D });

        Assert.True(selected.HasValue);
        Assert.Equal("same.png", selected.Candidate.Url);
        Assert.False(selected.UsesDefaultSource);
    }

    [Fact]
    public void OversizedSizesFallsBackWithoutScanningEveryToken() {
        string sizes = string.Join(" ", Enumerable.Repeat("invalid", 10_000));
        HtmlResponsiveImageSelection selected = HtmlResponsiveImageSelector.Select(
            "small.png 400w, large.png 800w", sizes, null,
            new HtmlResponsiveImageSelectionOptions {
                ViewportWidth = 800D,
                MaxSizesCharacters = 64
            });

        Assert.True(selected.HasValue);
        Assert.Equal("large.png", selected.Candidate.Url);
        Assert.Equal(800D, selected.SourceSize);
    }

    [Fact]
    public void ResourcePipelineUsesTheSharedSizesCharacterLimit() {
        HtmlConversionLimits limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxResponsiveImageSizesCharacters = 8;
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("""
            <img srcset="small.png 400w, large.png 800w" sizes="screen 400px">
            """, new HtmlResourcePipelineOptions {
                BaseUri = new Uri("https://example.test/"),
                Limits = limits,
                MediaWidth = 800D,
                MediaHeight = 600D
            });

        HtmlResourceReference image = Assert.Single(manifest.Resources,
            resource => resource.Kind == HtmlResourceKind.Image);
        Assert.Equal("https://example.test/large.png", image.ResolvedSource);
    }

    [Fact]
    public async Task RendererPropagatesAnExplicitSizesLimitAboveThePipelineDefault() {
        string sizes = new string(' ', 70_000) + "400px";
        HtmlConversionLimits limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxResponsiveImageSizesCharacters = 128 * 1024;
        HtmlConversionDocument document = HtmlConversionDocument.Parse($"""
            <img srcset="https://example.test/small.svg 400w, https://example.test/large.svg 800w"
                 sizes="{sizes}" width="20" height="20">
            """, new HtmlConversionDocumentOptions { Limits = limits });
        var requested = new List<string>();
        var options = new HtmlRenderOptions {
            ViewportWidth = 800D,
            ViewportHeight = 600D,
            ResourceResolver = (request, _) => {
                requested.Add(request.Uri.AbsoluteUri);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                    System.Text.Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='2' height='2'></svg>"),
                    "image/svg+xml"));
            }
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document, options);

        Assert.Contains("https://example.test/small.svg", requested);
        Assert.DoesNotContain("https://example.test/large.svg", requested);
        Assert.DoesNotContain(rendered.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    [Fact]
    public void SelectorNormalizesWidthDescriptorsAgainstTheFirstMatchingSize() {
        HtmlResponsiveImageSelection selected = HtmlResponsiveImageSelector.Select(
            "small.png 400w, medium.png 800w, large.png 1200w",
            "(max-width: 600px) 100vw, calc(50vw - 20px)", null,
            new HtmlResponsiveImageSelectionOptions {
                ViewportWidth = 800D,
                ViewportHeight = 600D,
                DevicePixelRatio = 2D
            });

        Assert.True(selected.HasValue);
        Assert.Equal("medium.png", selected.Candidate.Url);
        Assert.Equal(380D, selected.SourceSize, 5);
        Assert.Equal(800D / 380D, selected.EffectiveDensity, 5);
    }

    [Fact]
    public void SelectorUsesOneHundredViewportWidthWhenSizesIsUnsupported() {
        HtmlResponsiveImageSelection selected = HtmlResponsiveImageSelector.Select(
            "small.png 320w, large.png 640w", "50%", null,
            new HtmlResponsiveImageSelectionOptions { ViewportWidth = 500D, ViewportHeight = 400D });

        Assert.True(selected.HasValue);
        Assert.Equal("large.png", selected.Candidate.Url);
        Assert.Equal(500D, selected.SourceSize);
    }

    [Fact]
    public void SelectorRejectsMixedWidthAndDensityDescriptorSets() {
        HtmlResponsiveImageSelection selected = HtmlResponsiveImageSelector.Select(
            "small.png 400w, retina.png 2x", "100vw", "fallback.png");

        Assert.False(selected.HasValue);
    }

    [Fact]
    public void ResourceManifestWithMediaGeometryIncludesOnlyTheSelectedCandidate() {
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("""
            <img src="fallback.png"
                 srcset="small.png 400w, medium.png 800w, large.png 1200w"
                 sizes="(max-width: 600px) 100vw, 50vw">
            """, new HtmlResourcePipelineOptions {
                BaseUri = new Uri("https://example.test/"),
                ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile(),
                MediaWidth = 800D,
                MediaHeight = 600D,
                DevicePixelRatio = 2D
            });

        HtmlResourceReference image = Assert.Single(manifest.Resources,
            resource => resource.Kind == HtmlResourceKind.Image);
        Assert.Equal("https://example.test/medium.png", image.ResolvedSource);
    }

    [Fact]
    public void ResourceManifestFiltersRejectedCandidatesBeforeSelectingFallback() {
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("""
            <img src="https://example.test/fallback.png"
                 srcset="file:///blocked.png 1x, https://example.test/high.png 2x">
            """, new HtmlResourcePipelineOptions {
                BaseUri = new Uri("https://example.test/"),
                ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile(),
                MediaWidth = 800D,
                MediaHeight = 600D,
                DevicePixelRatio = 1D
            });

        HtmlResourceReference image = Assert.Single(manifest.Resources,
            resource => resource.Kind == HtmlResourceKind.Image);
        Assert.Equal("src", image.AttributeName);
        Assert.Equal("https://example.test/fallback.png", image.ResolvedSource);
        Assert.True(image.IsAllowed);
    }

    [Fact]
    public void PictureFallsBackWhenItsOnlyResponsiveCandidateIsRejected() {
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("""
            <picture>
              <source srcset="file:///blocked.png 1x">
              <img src="https://example.test/fallback.png">
            </picture>
            """, new HtmlResourcePipelineOptions {
                BaseUri = new Uri("https://example.test/"),
                ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile(),
                MediaWidth = 800D,
                MediaHeight = 600D
            });

        HtmlResourceReference image = Assert.Single(manifest.Resources,
            resource => resource.Kind == HtmlResourceKind.Image);
        Assert.Equal("img", image.ElementName);
        Assert.Equal("https://example.test/fallback.png", image.ResolvedSource);
    }

    [Fact]
    public void PictureSelectsAnAllowedCandidateAfterRejectedCandidatesAreRemoved() {
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("""
            <picture>
              <source srcset="file:///blocked.png 1x, https://example.test/high.png 2x">
              <img src="https://example.test/fallback.png">
            </picture>
            """, new HtmlResourcePipelineOptions {
                BaseUri = new Uri("https://example.test/"),
                ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile(),
                MediaWidth = 800D,
                MediaHeight = 600D,
                DevicePixelRatio = 1D
            });

        HtmlResourceReference image = Assert.Single(manifest.Resources,
            resource => resource.Kind == HtmlResourceKind.Image);
        Assert.Equal("source", image.ElementName);
        Assert.Equal("https://example.test/high.png", image.ResolvedSource);
    }

    [Fact]
    public void ResponsiveImagePreloadPlansOnlyTheSelectedCandidate() {
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("""
            <link rel="preload" as="image" href="fallback.png"
                  imagesrcset="one.png 1x, two.png 2x" imagesizes="100vw">
            """, new HtmlResourcePipelineOptions {
                BaseUri = new Uri("https://example.test/"),
                MediaWidth = 800D,
                MediaHeight = 600D,
                DevicePixelRatio = 2D
            });

        HtmlResourceReference image = Assert.Single(manifest.Resources,
            resource => resource.Kind == HtmlResourceKind.Image);
        Assert.Equal("imagesrcset", image.AttributeName);
        Assert.Equal("https://example.test/two.png", image.ResolvedSource);
    }

    [Fact]
    public void ManifestProvenanceKeepsAuthoredSrcSetWhenUrlMatchesSrc() {
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest("""
            <img src="same.png" srcset="same.png, high.png 2x">
            """, new HtmlResourcePipelineOptions {
                BaseUri = new Uri("https://example.test/"),
                MediaWidth = 800D,
                MediaHeight = 600D
            });

        HtmlResourceReference image = Assert.Single(manifest.Resources,
            resource => resource.Kind == HtmlResourceKind.Image);
        Assert.Equal("srcset", image.AttributeName);
        Assert.Equal("https://example.test/same.png", image.ResolvedSource);
    }
}
