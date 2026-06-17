using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlCoreTests {
    [Fact]
    public void HtmlDocumentParser_ResolvesBaseElementAgainstFallbackBaseUri() {
        var document = HtmlDocumentParser.ParseDocument("""<base href="images/"><p>Body</p>""");

        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(
            document,
            new Uri("https://example.test/articles/2026/"));

        Assert.Equal("https://example.test/articles/2026/images/", baseUri?.AbsoluteUri);
        Assert.Equal(document.Body, HtmlDocumentParser.GetConversionRoot(document, useBodyContentsOnly: true));
    }

    [Fact]
    public void HtmlDocumentParser_ResolvesProtocolRelativeBaseElementAgainstWebSchemeWhenFallbackIsFile() {
        var document = HtmlDocumentParser.ParseDocument("""<base href="//cdn.example.test/assets/"><img src="logo.png">""");

        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(
            document,
            new Uri("file:///C:/content/page.html"));

        Assert.Equal("https://cdn.example.test/assets/", baseUri?.AbsoluteUri);
    }

    [Fact]
    public void HtmlImageSourceResolver_ResolvesPictureSourceSetAgainstBaseUri() {
        var document = HtmlDocumentParser.ParseDocument("""
<picture>
  <source media="(min-width: 800px)" srcset="media/wide.webp 1x, media/wide@2x.webp 2x">
  <img src="media/fallback.png" alt="Storm">
</picture>
""");

        var picture = document.QuerySelector("picture")!;
        string source = HtmlImageSourceResolver.ResolvePictureSource(
            picture,
            new Uri("https://example.test/news/2026/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Equal("https://example.test/news/2026/media/wide.webp", source);

        string normalized = HtmlImageSourceResolver.ResolveNormalizedSrcSetAttributes(
            picture.QuerySelector("source")!,
            new Uri("https://example.test/news/2026/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile(),
            "srcset");

        Assert.Equal("https://example.test/news/2026/media/wide.webp 1x, https://example.test/news/2026/media/wide@2x.webp 2x", normalized);
    }

    [Fact]
    public void HtmlImageSourceResolver_OrdersParentPictureSourcesBeforeImageFallback() {
        var document = HtmlDocumentParser.ParseDocument("""
<picture>
  <source srcset="media/hero.webp 1x">
  <img src="media/fallback.png" alt="Hero">
</picture>
""");

        var image = document.QuerySelector("img")!;
        IReadOnlyList<string> candidates = HtmlImageSourceResolver.ResolveImageSourceCandidates(
            image,
            new Uri("https://example.test/news/2026/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Collection(
            candidates,
            source => Assert.Equal("https://example.test/news/2026/media/hero.webp", source),
            source => Assert.Equal("https://example.test/news/2026/media/fallback.png", source));

        string resolved = HtmlImageSourceResolver.ResolveImageSource(
            image,
            new Uri("https://example.test/news/2026/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Equal("https://example.test/news/2026/media/hero.webp", resolved);
    }

    [Fact]
    public void HtmlImageSourceResolver_UsesLazyAttributesBeforePlaceholderSource() {
        var document = HtmlDocumentParser.ParseDocument("""<img src="data:image/gif;base64,AAAA" data-lazy-src="media/photo.png">""");
        string source = HtmlImageSourceResolver.ResolveImageSource(
            document.QuerySelector("img")!,
            new Uri("https://example.test/gallery/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Equal("https://example.test/gallery/media/photo.png", source);
    }

    [Fact]
    public void HtmlImageSourceResolver_UsesSourceSetBeforeImageSourceFallback() {
        var document = HtmlDocumentParser.ParseDocument("""<img src="media/fallback.png" srcset="media/hero.webp 1x" alt="Hero">""");
        var image = document.QuerySelector("img")!;

        IReadOnlyList<string> candidates = HtmlImageSourceResolver.ResolveImageSourceCandidates(
            image,
            new Uri("https://example.test/gallery/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Collection(
            candidates,
            source => Assert.Equal("https://example.test/gallery/media/hero.webp", source),
            source => Assert.Equal("https://example.test/gallery/media/fallback.png", source));
        Assert.Equal(
            "https://example.test/gallery/media/hero.webp",
            HtmlImageSourceResolver.ResolveImageSource(image, new Uri("https://example.test/gallery/"), HtmlUrlPolicy.CreateOfficeIMOProfile()));
    }

    [Fact]
    public void HtmlImageSourceResolver_LimitsResponsiveCandidatesAndKeepsImageFallback() {
        var document = HtmlDocumentParser.ParseDocument("""
<picture>
  <source srcset="media/one.webp 1x, media/one.webp 2x, media/two.webp 3x, media/three.webp 4x">
  <img src="media/fallback.png" srcset="media/four.webp 4x" alt="Hero">
</picture>
""");
        var image = document.QuerySelector("img")!;

        IReadOnlyList<string> candidates = HtmlImageSourceResolver.ResolveImageSourceCandidates(
            image,
            new Uri("https://example.test/gallery/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile(),
            allowParentPictureFallback: true,
            maxResponsiveCandidates: 2);

        Assert.Collection(
            candidates,
            source => Assert.Equal("https://example.test/gallery/media/one.webp", source),
            source => Assert.Equal("https://example.test/gallery/media/two.webp", source),
            source => Assert.Equal("https://example.test/gallery/media/fallback.png", source));
    }

    [Fact]
    public void HtmlImageSourceResolver_CountsRejectedSrcSetEntriesTowardExpansionLimit() {
        var document = HtmlDocumentParser.ParseDocument("""
<img srcset="javascript:alert(1) 1x, javascript:alert(2) 2x, media/good.webp 3x" src="media/fallback.png" alt="Hero">
""");
        var image = document.QuerySelector("img")!;

        IReadOnlyList<string> candidates = HtmlImageSourceResolver.ResolveImageSourceCandidates(
            image,
            new Uri("https://example.test/gallery/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile(),
            allowParentPictureFallback: true,
            maxResponsiveCandidates: 2);

        var source = Assert.Single(candidates);
        Assert.Equal("https://example.test/gallery/media/fallback.png", source);
    }

    [Fact]
    public void HtmlSrcSetParser_CanLimitCandidateExpansion() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse(
            "one.png 1x, two.png 2x, three.png 3x",
            maxCandidates: 2);

        Assert.Collection(
            candidates,
            candidate => Assert.Equal("one.png", candidate.Url),
            candidate => Assert.Equal("two.png", candidate.Url));
    }

    [Fact]
    public void HtmlSrcSetParser_SplitsCommaSeparatedCandidatesWithoutWhitespace() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("small.png,large.png 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("small.png", candidate.Url);
                Assert.Equal(string.Empty, candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("large.png", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });

        string normalized = HtmlImageSourceResolver.ResolveNormalizedSrcSet(
            "small.png,large.png 2x",
            new Uri("https://example.test/images/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Equal("https://example.test/images/small.png, https://example.test/images/large.png 2x", normalized);
    }

    [Fact]
    public void HtmlSrcSetParser_SplitsQueryCandidatesAtCommaSeparators() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("small.png?v=1,large.png?v=1 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("small.png?v=1", candidate.Url);
                Assert.Equal(string.Empty, candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("large.png?v=1", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });
    }

    [Fact]
    public void HtmlSrcSetParser_SplitsBareExtensionlessCandidatesAtCommaSeparators() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("small,large 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("small", candidate.Url);
                Assert.Equal(string.Empty, candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("large", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });
    }

    [Fact]
    public void HtmlSrcSetParser_SplitsExtensionlessQueryCandidatesAtCommaSeparators() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("image?w=200,image?w=400 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("image?w=200", candidate.Url);
                Assert.Equal(string.Empty, candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("image?w=400", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });
    }

    [Fact]
    public void HtmlSrcSetParser_PreservesCommaValuedQueryStringsBeforeFollowingCandidate() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("photo.jpg?tags=red,blue 1x, photo@2x.jpg 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("photo.jpg?tags=red,blue", candidate.Url);
                Assert.Equal("1x", candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("photo@2x.jpg", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });
    }

    [Fact]
    public void HtmlSrcSetParser_SplitsDataUriCandidateBeforeFollowingUrl() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse("data:image/png;base64,AAAA, https://cdn.test/fallback.png 2x");

        Assert.Collection(
            candidates,
            candidate => {
                Assert.Equal("data:image/png;base64,AAAA", candidate.Url);
                Assert.Equal(string.Empty, candidate.Descriptor);
            },
            candidate => {
                Assert.Equal("https://cdn.test/fallback.png", candidate.Url);
                Assert.Equal("2x", candidate.Descriptor);
            });
    }

    [Fact]
    public void HtmlImageDataUri_ParsesAndDecodesBase64Images() {
        string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\"/>";
        string dataUri = "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));

        Assert.True(HtmlImageDataUri.TryParse(dataUri, out var image));
        Assert.True(image.IsBase64);
        Assert.Equal("image/svg+xml", image.MediaType);
        Assert.Equal(".svg", image.FileExtension);
        Assert.Equal(svg, image.DecodeText());
        Assert.Equal(Encoding.UTF8.GetByteCount(svg), image.EstimateDecodedByteCount());
    }

    [Fact]
    public void HtmlImageDataUri_TryDecodeBytesReturnsFalseForBadEscapes() {
        Assert.True(HtmlImageDataUri.TryParse("data:image/png;base64,%ZZ", out var image));
        Assert.True(image.IsBase64);
        Assert.False(image.TryDecodeBytes(out byte[] bytes));
        Assert.Empty(bytes);
    }

    [Fact]
    public void HtmlImageDataUri_DecodesNonBase64PercentEscapesAsBytes() {
        Assert.True(HtmlImageDataUri.TryParse("data:image/png,%89PNG%0D%0A%1A%0A", out var image));

        Assert.False(image.IsBase64);
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, image.DecodeBytes());
        Assert.Equal(8, image.EstimateDecodedByteCount());
    }

    [Fact]
    public void HtmlImageDataUri_MatchesOnlyExactBase64Flag() {
        string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\"/>";
        string dataUri = "data:image/svg+xml;name=base64," + Uri.EscapeDataString(svg);

        Assert.True(HtmlImageDataUri.TryParse(dataUri, out var image));
        Assert.False(image.IsBase64);
        Assert.Equal(svg, image.DecodeText());
    }

    [Fact]
    public void HtmlImageDataUri_IgnoresBase64WhitespaceWhenEstimatingDecodedBytes() {
        Assert.True(HtmlImageDataUri.TryParse("data:image/png;base64,QUJD%0A", out var image));

        Assert.True(image.IsBase64);
        Assert.Equal(3, image.EstimateDecodedByteCount());
        Assert.Equal(new byte[] { 65, 66, 67 }, image.DecodeBytes());
    }

    [Fact]
    public void HtmlUrlPolicyEvaluator_RejectsScriptUrlsAndCanResolveRelativeUrls() {
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();

        string rejected = HtmlUrlPolicyEvaluator.ResolveUrl("javascript:alert(1)", new Uri("https://example.test/"), policy);
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl("../docs/index.html", new Uri("https://example.test/news/2026/"), policy);
        string rootRelative = HtmlUrlPolicyEvaluator.ResolveUrl("/img/demo.png", new Uri("https://example.test/news/2026/"), policy);

        Assert.Equal(string.Empty, rejected);
        Assert.Equal("https://example.test/news/docs/index.html", resolved);
        Assert.Equal("https://example.test/img/demo.png", rootRelative);
    }

    [Fact]
    public void HtmlUrlPolicyEvaluator_ResolvesProtocolRelativeUrlsAgainstWebSchemeWhenBaseIsFile() {
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();

        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl("//cdn.example.test/app.png", new Uri("file:///C:/content/page.html"), policy);

        Assert.Equal("https://cdn.example.test/app.png", resolved);
    }

    [Fact]
    public void HtmlUrlPolicyEvaluator_ResolvesProtocolRelativeUrlsWithoutBaseAgainstHttps() {
        var policy = HtmlUrlPolicy.CreateOfficeIMOProfile();

        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl("//cdn.example.test/app.png", null, policy);

        Assert.Equal("https://cdn.example.test/app.png", resolved);
    }

    [Theory]
    [InlineData("java\nscript:alert(1)")]
    [InlineData("vb\rscript:msgbox(1)")]
    [InlineData("java\tscript:alert(1)")]
    public void HtmlUrlPolicyEvaluator_RejectsUrlsWithEmbeddedControlCharacters(string rawUrl) {
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();

        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed(rawUrl, policy));
        Assert.Equal(string.Empty, HtmlUrlPolicyEvaluator.ResolveUrl(rawUrl, new Uri("https://example.test/"), policy));
    }

    [Theory]
    [InlineData("C:secret.docx")]
    [InlineData("C:\\secret.docx")]
    [InlineData("C:/secret.docx")]
    public void HtmlUrlPolicyEvaluator_RejectsWindowsDrivePathsWhenFileUrlsAreDisallowed(string rawUrl) {
        var policy = HtmlUrlPolicy.CreateHyperlinkProfile();

        Assert.False(HtmlUrlPolicyEvaluator.IsAllowed(rawUrl, policy));
        Assert.Equal(string.Empty, HtmlUrlPolicyEvaluator.ResolveUrl(rawUrl, new Uri("https://example.test/"), policy));
    }
}
