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
}
