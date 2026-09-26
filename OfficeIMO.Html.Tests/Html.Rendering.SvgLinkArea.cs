using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPdf_OuterLinkCoversInlineSvgViewport() {
        const string html = "<a href='https://example.test/logo'><svg xmlns='http://www.w3.org/2000/svg' "
            + "width='46' height='46' viewBox='0 0 360 360'><g><path d='M2 2h100v100H2z' fill='blue'/>"
            + "<path d='M240 240h100v100H240z' fill='blue'/></g></svg></a>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes();
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(
            PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri("https://example.test/logo"));

        Assert.InRange(link.Width, 34D, 35D);
        Assert.InRange(link.Height, 34D, 35D);
    }

    [Fact]
    public void HtmlPdf_OuterFragmentLinkCoversInlineSvgViewport() {
        const string html = "<a href='#details'><svg xmlns='http://www.w3.org/2000/svg' "
            + "width='46' height='46' viewBox='0 0 360 360'><g>"
            + "<path d='M2 2h100v100H2z' fill='blue'/></g></svg></a>"
            + "<p id='details'>Details</p>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Contains("html-fragment:details", info.LinkDestinationNames);
        Assert.DoesNotContain("#details", info.LinkUris);
    }

    [Fact]
    public void HtmlPdf_FlexAnchorLinkCoversItsPaddedBorderBox() {
        const string html = "<a href='https://example.test/logo' style='display:flex;width:100px;"
            + "padding:8px 0;margin:0;align-items:center'>"
            + "<svg xmlns='http://www.w3.org/2000/svg' width='46' height='46'>"
            + "<rect width='46' height='46' fill='blue'/></svg></a>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            Margins = HtmlRenderMargins.All(0D)
        });
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(
            PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri("https://example.test/logo"));

        Assert.InRange(link.Width, 74.9D, 75.1D);
        Assert.InRange(link.Height, 46.4D, 46.6D);
    }

    [Fact]
    public void HtmlPdf_BlockAnchorsWithTheSameTargetKeepSeparatePaddedAreas() {
        const string html = "<a href='https://example.test/target' style='display:block;width:100px;"
            + "height:20px;padding:11px 0;margin:0'>First</a>"
            + "<a href='https://example.test/target' style='display:block;width:100px;"
            + "height:20px;padding:11px 0;margin:20px 0 0'>Second</a>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<PdfCore.PdfLogicalLinkAnnotation> links = PdfCore.PdfDocumentReadResult
            .Load(pdf).GetLinksByUri("https://example.test/target");

        Assert.Equal(2, links.Count);
        Assert.All(links, link => {
            Assert.InRange(link.Width, 74.9D, 75.1D);
            Assert.InRange(link.Height, 31.4D, 31.6D);
        });
    }

    [Fact]
    public void HtmlPdf_ClippedBlockAnchorRetainsVisibleLink() {
        const string html = "<div style='clip-path:polygon(0 0,100% 0,0 100%);width:100px;height:40px'>"
            + "<a href='https://example.test/inside' style='display:block;width:60px;height:20px'>Inside</a>"
            + "</div>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf)
            .GetLinksByUri("https://example.test/inside"));
    }

    [Fact]
    public void HtmlPdf_ClippedInlineAnchorDoesNotDuplicateItsLink() {
        const string html = "<div style='clip-path:polygon(0 0,100% 0,0 100%);width:100px;height:40px'>"
            + "<a href='https://example.test/inside'>Inside</a></div>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf)
            .GetLinksByUri("https://example.test/inside"));
    }

    [Fact]
    public void HtmlPdf_InlineAnchorLinkIncludesPaintedPadding() {
        const string html = "<div style='margin:30px 0 0 30px'><a href='https://example.test/padded' style='font:16px Arial;"
            + "padding:10px 20px;background:yellow'>Text</a></div>";
        var options = new HtmlRenderOptions { Margins = HtmlRenderMargins.All(0D) };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderAnchorFragment fragment = Assert.Single(rendered.Pages[0].Visuals
            .OfType<HtmlRenderAnchorFragment>());

        Assert.InRange(fragment.Width, 70D, 100D);
        Assert.InRange(fragment.Height, 35D, 50D);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(options));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(PdfCore.PdfDocumentReadResult
            .Load(pdf).GetLinksByUri("https://example.test/padded"));
        Assert.InRange(link.Width, 50D, 60D);
        Assert.InRange(link.Height, 25D, 40D);
        Assert.Equal("Text", link.Contents);
    }
}
