using OfficeIMO.Html;
using System.Threading;

namespace OfficeIMO.Chm.Tests;

public sealed class ProjectionReferenceTests {
    [Theory]
    [InlineData("<svg xmlns='http://www.w3.org/2000/svg'><image href='views.svg#closeup'/></svg>", "image", "href")]
    [InlineData("<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'><image xlink:href='views.svg#closeup'/></svg>", "image", "xlink:href")]
    [InlineData("<svg xmlns='http://www.w3.org/2000/svg'><defs><filter id='view'><feImage href='views.svg#closeup'/></filter></defs></svg>", "feImage", "href")]
    [InlineData("<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'><defs><filter id='view'><feImage xlink:href='views.svg#closeup'/></filter></defs></svg>", "feImage", "xlink:href")]
    [InlineData("<img src='views.svg#closeup' alt='Close view'>", "img", "src")]
    [InlineData("<img data-src='views.svg#closeup' alt='Close view'>", "img", "data-src")]
    [InlineData("<picture><source srcset='views.svg#closeup 1x'><img src='views.svg#closeup' alt='Close view'></picture>", "source", "srcset")]
    public void EmbeddedSvgViewsRetainTheirResourceFragment(string markup, string tag, string attribute) {
        byte[] svg = ChmFixture.Html("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 10'><view id='closeup' viewBox='10 0 10 10'/><rect width='20' height='10'/></svg>");
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/topic.html"] = ChmFixture.Html(markup), ["/views.svg"] = svg
        }));
        var projection = book.ToHtmlDocumentResult();
        string value = Assert.Single(projection.Value.Document.QuerySelector(tag)!.Attributes, item => item.Name == attribute).Value;
        if (attribute == "srcset") value = Assert.Single(HtmlSrcSetParser.Parse(value)).Url;
        Assert.True(HtmlImageDataUri.TryParse(value, out HtmlImageDataUri parsed));
        Assert.Equal("#closeup", parsed.Fragment);
        Assert.Equal(svg, parsed.DecodeBytes());
    }

    [Theory]
    [InlineData("href", "#part", "#chm-topic-1-part")]
    [InlineData("xlink:href", "#part", "#chm-topic-1-part")]
    [InlineData("href", "second.html#part", "#chm-topic-2-part")]
    [InlineData("xlink:href", "second.html#part", "#chm-topic-2-part")]
    public void SvgHyperlinksAreRewrittenOnce(string attribute, string target, string expected) {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/first.html"] = ChmFixture.Html("<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>" +
                "<path id='part'/><a " + attribute + "='" + target + "'>Continue</a></svg>"),
            ["/second.html"] = ChmFixture.Html("<p id='part'>Destination</p>")
        }));
        var projection = book.ToHtmlDocumentResult();
        Assert.Equal(expected, Assert.Single(projection.Value.Document.QuerySelector("a")!.Attributes, item => item.LocalName == "href").Value);
    }

    [Theory]
    [InlineData("image", "href", true)]
    [InlineData("image", "xlink:href", true)]
    [InlineData("image", "href", false)]
    [InlineData("image", "xlink:href", false)]
    [InlineData("feImage", "href", true)]
    [InlineData("feImage", "xlink:href", true)]
    [InlineData("feImage", "href", false)]
    [InlineData("feImage", "xlink:href", false)]
    public void SvgImagesRetainArchiveResources(string tag, string attribute, bool embed) {
        ChmDocument book = ImageBook("<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>" +
            "<" + tag + " " + attribute + "='pixel.png'/></svg>");
        var projection = book.ToHtmlDocumentResult(new ChmConversionOptions { EmbedImages = embed });
        var reference = Assert.Single(projection.Value.Document.QuerySelector(tag)!.Attributes, item => item.LocalName == "href");
        Assert.Equal(attribute.StartsWith("xlink:", StringComparison.Ordinal) ? "http://www.w3.org/1999/xlink" : string.Empty, reference.NamespaceUri);
        if (embed) Assert.Equal("data:image/png;base64," + Convert.ToBase64String(book.FindEntry("/pixel.png")!.GetBytes()), reference.Value);
        else Assert.Equal("chm://archive/pixel.png", reference.Value);
    }

    [Theory]
    [InlineData("href")]
    [InlineData("xlink:href")]
    public void MissingSvgImagesProduceResourceOmissions(string attribute) {
        var projection = ImageBook("<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'><image " +
            attribute + "='missing.png'/></svg>").ToHtmlDocumentResult();
        Assert.DoesNotContain(projection.Value.Document.QuerySelector("image")!.Attributes, item => item.LocalName == "href");
        Assert.Contains(projection.Report.FidelityDiagnostics, item => item.Code == "CHM_IMAGE_MISSING");
    }

    [Fact]
    public void SvgImageReferencesConsumeTheSharedEmbeddingBudget() {
        ChmDocument book = ImageBook("<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>" +
            "<image href='pixel.png'/><image xlink:href='pixel.png'/></svg>");
        Assert.Equal("CHM_CONVERSION_LIMIT", Assert.Throws<ChmReadException>(() => book.ToHtmlDocumentResult(
            new ChmConversionOptions { MaxTotalEmbeddedImageBytes = book.FindEntry("/pixel.png")!.Length })).Code);
    }

    [Fact]
    public void SvgImagesPreserveExternalDataAndSameTopicFragmentReferences() {
        const string data = "data:image/png;base64,AQID";
        var projection = ImageBook("<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>" +
            "<path id='shape'/><image href='https://example.invalid/pixel.png'/><image xlink:href='" + data + "'/><image href='#shape'/></svg>").ToHtmlDocumentResult();
        var images = projection.Value.Document.QuerySelectorAll("image");
        Assert.Equal("https://example.invalid/pixel.png", images[0].GetAttribute("href"));
        Assert.Equal(data, Assert.Single(images[1].Attributes, item => item.LocalName == "href").Value);
        Assert.Equal("#chm-topic-1-shape", images[2].GetAttribute("href"));
    }

    [Fact]
    public void ExternalSvgUseDefinitionsRetainTheirArchiveResourceUri() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/topic.html"] = ChmFixture.Html("<svg xmlns='http://www.w3.org/2000/svg'><use href='icons.svg#shape'/></svg>"),
            ["/icons.svg"] = ChmFixture.Html("<svg xmlns='http://www.w3.org/2000/svg'><path id='shape'/></svg>")
        }));
        var projection = book.ToHtmlDocumentResult();
        Assert.Equal("chm://archive/icons.svg#shape", projection.Value.Document.QuerySelector("use")!.GetAttribute("href"));
    }

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
