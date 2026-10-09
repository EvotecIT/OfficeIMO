using OfficeIMO.Html;
using System.Threading;

namespace OfficeIMO.Chm.Tests;

public sealed class ProjectionReferenceTests {
    [Theory]
    [InlineData("<img srcset='pixel.png 1x, large.png 2x' alt='Pixel'>")]
    [InlineData("<picture><source srcset='large.png 2x'><img src='pixel.png' alt='Pixel'></picture>")]
    [InlineData("<img data-src='pixel.png' alt='Pixel'>")]
    [InlineData("<img data-original='pixel.png' alt='Pixel'>")]
    [InlineData("<img data-original-src='pixel.png' alt='Pixel'>")]
    [InlineData("<img data-lazy-src='pixel.png' alt='Pixel'>")]
    [InlineData("<img data-srcset='pixel.png 1x, large.png 2x' alt='Pixel'>")]
    [InlineData("<img data-original-srcset='pixel.png 1x, large.png 2x' alt='Pixel'>")]
    [InlineData("<img data-lazy-srcset='pixel.png 1x, large.png 2x' alt='Pixel'>")]
    public void ResponsiveAndLazyImagesArePortableInMarkdown(string image) {
        ChmDocument book = ImageBook(image);
        var markdown = book.ToMarkdownResult();
        Assert.Contains("data:image/png;base64,", markdown.Value);
        Assert.DoesNotContain("chm://", markdown.Value);
        var html = book.ToHtmlDocumentResult();
        foreach (var element in html.Value.Document.QuerySelectorAll("img,source"))
            Assert.DoesNotContain(element.Attributes, attribute => attribute.Value.Contains("chm://"));
        string? srcset = html.Value.Document.QuerySelector("img")!.GetAttribute("srcset");
        if (srcset != null) Assert.Equal(new[] { "1x", "2x" }, HtmlSrcSetParser.Parse(srcset).Select(candidate => candidate.Descriptor));
    }

    [Fact]
    public void EachEmbeddedCandidateConsumesTheAggregateImageBudget() {
        ChmDocument book = ImageBook("<img srcset='pixel.png 1x, large.png 2x' alt='Pixel'>");
        long bytes = book.FindEntry("/pixel.png")!.Length;
        Assert.Equal("CHM_CONVERSION_LIMIT", Assert.Throws<ChmReadException>(() =>
            book.ToHtmlDocumentResult(new ChmConversionOptions { MaxTotalEmbeddedImageBytes = bytes })).Code);
        var external = ImageBook("<img srcset='https://example.invalid/pixel.png 1x' alt='Pixel'>").ToHtmlDocumentResult();
        Assert.Equal("https://example.invalid/pixel.png", Assert.Single(HtmlSrcSetParser.Parse(
            external.Value.Document.QuerySelector("img")!.GetAttribute("srcset"))).Url);
    }

    [Fact]
    public void InlineEventAttributesAloneProduceAnActiveContentOmission() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/topic.html"] = ChmFixture.Html("<p onclick='danger()'>Visible text</p>")
        }));
        var projection = book.ToHtmlDocumentResult();
        Assert.Contains("Visible text", projection.Value.SourceHtml);
        Assert.DoesNotContain("onclick", projection.Value.SourceHtml);
        Assert.Contains(projection.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "CHM_ACTIVE_CONTENT_OMITTED" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Throws<OfficeConversionException>(() => projection.RequireNoLoss());
    }

    [Fact]
    public void SvgFragmentLinksAndLocalPaintReferencesFollowPrefixedIds() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/topic.html"] = ChmFixture.Html("<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>" +
                "<defs><path id='shape' d='M0 0L1 1'/><linearGradient id='paint'/><clipPath id='clip'/><mask id='mask'/><marker id='marker'/></defs>" +
                "<use href='#shape'/><use xlink:href='#shape'/><rect fill='url(#paint)' clip-path='url(\"#clip\")' mask='url(#mask)' " +
                "marker-start='url(#marker)' style='stroke:url(#paint);filter:url(https://example.invalid/filter)'/></svg>")
        }));
        var projection = book.ToHtmlDocumentResult();
        var uses = projection.Value.Document.QuerySelectorAll("use");
        Assert.Equal("#chm-topic-1-shape", Assert.Single(uses[0].Attributes, attribute => attribute.LocalName == "href").Value);
        var xlink = Assert.Single(uses[1].Attributes, attribute => attribute.LocalName == "href");
        Assert.Equal("http://www.w3.org/1999/xlink", xlink.NamespaceUri);
        Assert.Equal("#chm-topic-1-shape", xlink.Value);
        var rectangle = projection.Value.Document.QuerySelector("rect")!;
        Assert.Equal("url(#chm-topic-1-paint)", rectangle.GetAttribute("fill"));
        Assert.Equal("url(\"#chm-topic-1-clip\")", rectangle.GetAttribute("clip-path"));
        Assert.Equal("url(#chm-topic-1-mask)", rectangle.GetAttribute("mask"));
        Assert.Equal("url(#chm-topic-1-marker)", rectangle.GetAttribute("marker-start"));
        Assert.Contains("#chm-topic-1-paint", rectangle.GetAttribute("style"));
        Assert.Contains("https://example.invalid/filter", rectangle.GetAttribute("style"));
    }

    private static ChmDocument ImageBook(string image) {
        byte[] pixel = ChmDocument.Load(ChmFixture.Book()).FindEntry("/pixel.png")!.GetBytes();
        return ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/topic.html"] = ChmFixture.Html(image), ["/pixel.png"] = pixel, ["/large.png"] = pixel
        }));
    }
}
