using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("", true)]
    [InlineData("min-width:calc(50% + 40px)", false)]
    public void HtmlPdf_InlineFlexRetainsLinkAcrossOverflowingIcon(string constraints, bool contained) {
        const string uri = "https://example.test/next";
        string html = "<style>body{margin:0;font:20px/20px Arial}</style>"
            + "<a href='" + uri + "' style='display:inline-flex;align-items:center'>"
            + "<span style='margin-right:8px;" + constraints + "'>Next Page</span>"
            + "<svg id='icon' width='20' height='20' viewBox='0 0 20 20' style='flex-shrink:0'"
            + " xmlns='http://www.w3.org/2000/svg'><circle cx='10' cy='10' r='10' fill='red'/></svg></a>";
        var options = new HtmlToPdfOptions { ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) };
        HtmlPdfRenderRequestResult result = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));
        HtmlRenderDocument rendered = result.RenderResult.Document;
        HtmlRenderDrawing icon = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        byte[] pdf = result.ToBytes();
        IReadOnlyList<PdfCore.PdfLogicalLinkAnnotation> links = PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(uri);

        if (contained) Assert.Single(links);
        Assert.Contains(links, link => link.X1 <= icon.X * .75D + .01D
            && link.X2 >= (icon.X + icon.Width) * .75D - .01D);
    }

    [Theory]
    [InlineData(19.5D, 20D)]
    [InlineData(20D, 19.5D)]
    public void HtmlPdf_FractionalSvgOverflowRetainsClickablePaint(double anchorWidth, double anchorHeight) {
        const string uri = "https://example.test/fractional";
        string html = $"<style>body{{margin:0}}</style><a href='{uri}' "
            + $"style='display:inline-block;width:{anchorWidth.ToString(System.Globalization.CultureInfo.InvariantCulture)}px;"
            + $"height:{anchorHeight.ToString(System.Globalization.CultureInfo.InvariantCulture)}px'>"
            + "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20'>"
            + "<rect width='20' height='20' fill='red'/></svg></a>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<PdfCore.PdfLogicalLinkAnnotation> links = PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(uri);
        Assert.Contains(links, link => link.Width >= 14.99D && link.Height >= 14.99D);
    }

    [Theory]
    [InlineData("")]
    [InlineData("position:relative;left:30px;top:40px")]
    public void HtmlPdf_TextOverflowKeepsLinkBeyondTheAnchorBox(string position) {
        const string uri = "https://example.test/text-overflow";
        byte[] pdf = HtmlConversionDocument.Parse("<style>body{margin:0;font:24px/32px Arial}</style>"
            + "<a href='" + uri + "' style='display:inline-block;width:20px;" + position + "'>NORTHWIND</a>")
            .ToPdfBytes(new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D) });
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf);
        var letters = parsed.GetPage(1).Letters;
        Assert.Equal("NORTHWIND", string.Concat(letters.Select(letter => letter.Value)));
        IReadOnlyList<PdfCore.PdfLogicalLinkAnnotation> links = PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(uri);
        Assert.Contains(links, link => link.X1 <= letters[0].StartBaseLine.X + .01D
            && link.X2 >= letters[letters.Count - 1].EndBaseLine.X - .01D);
    }

    [Theory]
    [InlineData("")]
    [InlineData("position:relative;left:30px;top:40px")]
    public void HtmlPdf_NegativeLetterSpacingRetainsPaintedGlyphAndPositionedLinks(string position) {
        const string uri = "https://example.test/negative-spacing";
        byte[] pdf = HtmlConversionDocument.Parse("<style>body{margin:0;font:24px/32px Arial}</style>"
            + "<a href='" + uri + "' style='display:inline-block;width:10px;letter-spacing:-12px;" + position + "'>A</a>")
            .ToPdfBytes(new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D) });
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf);
        var letter = Assert.Single(parsed.GetPage(1).Letters);
        IReadOnlyList<PdfCore.PdfLogicalLinkAnnotation> links = PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(uri);
        Assert.Contains(links, link => link.X1 <= letter.StartBaseLine.X + .01D
            && link.X2 >= letter.EndBaseLine.X - .01D);
        Assert.All(links, link => Assert.True(link.X1 >= letter.StartBaseLine.X - .01D));
    }

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

    [Fact]
    public void HtmlPdf_InlineAnchorUsesFaceHeightInsteadOfExtraLineLeading() {
        const string html = "<div style='font:16px/24px Example'>"
            + "<a href='https://example.test/inline'>Link text</a></div>";
        var options = new HtmlRenderOptions {
            Margins = HtmlRenderMargins.All(0D),
            FallbackTextFaceMetrics = (_, _, _) => new HtmlTextFaceMetrics(19D, 15D)
        };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderAnchorFragment fragment = Assert.Single(rendered.Pages[0].Visuals
            .OfType<HtmlRenderAnchorFragment>());
        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>());

        Assert.Equal(19D, fragment.Height, 3);
        Assert.Equal(text.Y, fragment.Y, 3);
        byte[] pdf = HtmlPdfRenderedConverter.CreatePdf(rendered, new HtmlToPdfOptions(options),
            System.Threading.CancellationToken.None).Document.ToBytes();
        Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri("https://example.test/inline"));
    }
}
