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
    public void HtmlImageSourceResolver_UsesLazyAttributesBeforePlaceholderSource() {
        var document = HtmlDocumentParser.ParseDocument("""<img src="data:image/gif;base64,AAAA" data-lazy-src="media/photo.png">""");
        string source = HtmlImageSourceResolver.ResolveImageSource(
            document.QuerySelector("img")!,
            new Uri("https://example.test/gallery/"),
            HtmlUrlPolicy.CreateOfficeIMOProfile());

        Assert.Equal("https://example.test/gallery/media/photo.png", source);
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
    public void HtmlUrlPolicyEvaluator_RejectsScriptUrlsAndCanResolveRelativeUrls() {
        var policy = HtmlUrlPolicy.CreateWebOnlyProfile();

        string rejected = HtmlUrlPolicyEvaluator.ResolveUrl("javascript:alert(1)", new Uri("https://example.test/"), policy);
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl("../docs/index.html", new Uri("https://example.test/news/2026/"), policy);
        string rootRelative = HtmlUrlPolicyEvaluator.ResolveUrl("/img/demo.png", new Uri("https://example.test/news/2026/"), policy);

        Assert.Equal(string.Empty, rejected);
        Assert.Equal("https://example.test/news/docs/index.html", resolved);
        Assert.Equal("https://example.test/img/demo.png", rootRelative);
    }
}
