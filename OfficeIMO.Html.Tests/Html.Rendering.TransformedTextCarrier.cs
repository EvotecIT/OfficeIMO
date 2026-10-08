using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("translateX(-300px)", false)]
    [InlineData("scale(.8) translateX(-300px)", false)]
    [InlineData("rotate(8deg) translateX(-300px)", false)]
    [InlineData("translateX(-300px)", true)]
    public void TransformedLogicalText_UsesTheVisiblePhysicalPageForItsCarrier(string transform, bool clipped) {
        string content = "<div style='position:relative;top:50px;width:200px;white-space:nowrap;"
            + "font:20px/30px Courier;letter-spacing:1px;transform-origin:0 0;transform:" + transform + "'>"
            + new string('X', 30) + " TransformCarrierMarker TrailingVisibleMarker</div>";
        if (clipped) content = "<div style='overflow:hidden;width:350px;height:110px'>" + content + "</div>";
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0}</style>"
            + content + "<p style='margin-top:130px'>SiblingMarker</p>";
        HtmlPdfRenderRequestResult result = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf,
                new HtmlToPdfOptions { ViewportWidth = 400D, ViewportHeight = 400D,
                    Margins = HtmlRenderMargins.All(0D), ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() }));
        PdfReadPage page = Assert.Single(PdfReadDocument.Open(result.ToBytes()).Pages);
        PdfTextSpan carrier = Assert.Single(page.GetTextSpans(), span => span.Text.Contains("TransformCarrierMarker", StringComparison.Ordinal));
        AssertCarrierWithinPhysicalPage(page, carrier);
        Assert.Single(page.GetTextSpans(), span => span.Text == "SiblingMarker");
    }

    [Fact]
    public void TransformedLogicalText_WideAuthoredPathClipRetainsPhysicalPageWindow() {
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0}</style>"
            + "<div style='clip-path:inset(0);position:relative;left:-100px;width:600px;height:180px'>"
            + "<div style='position:relative;top:50px;width:200px;white-space:nowrap;"
            + "font:20px/30px Courier;letter-spacing:1px;transform-origin:0 0;transform:translateX(-300px)'>"
            + new string('X', 30) + " TransformCarrierMarker TrailingVisibleMarker</div></div>";
        PdfReadPage page = Assert.Single(RenderTransformedCarrier(html).Pages);
        PdfTextSpan carrier = Assert.Single(page.GetTextSpans(), span => span.Text.Contains("TransformCarrierMarker", StringComparison.Ordinal));
        AssertCarrierWithinPhysicalPage(page, carrier);
    }

    [Fact]
    public void TransformedLogicalText_CornerIntersectionKeepsAnInPageCarrier() {
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0}</style>"
            + "<div style='width:400px;height:400px;overflow:hidden'>"
            + "<div style='width:400px;transform:matrix(1,1,-1,1,0,0);transform-origin:0 0'>"
            + "<div style='padding-top:189px;white-space:nowrap;font:12px/12px Courier;letter-spacing:1px'>"
            + new string('X', 120) + " TransformCarrierMarker</div></div></div>";
        PdfReadPage page = Assert.Single(RenderTransformedCarrier(html).Pages);
        PdfTextSpan carrier = Assert.Single(page.GetTextSpans(), span => span.Text.Contains("TransformCarrierMarker", StringComparison.Ordinal));
        AssertCarrierWithinPhysicalPage(page, carrier);
    }

    private static PdfReadDocument RenderTransformedCarrier(string html) => PdfReadDocument.Open(
        HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf,
                new HtmlToPdfOptions { ViewportWidth = 400D, ViewportHeight = 400D,
                    Margins = HtmlRenderMargins.All(0D), ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() })).ToBytes());

    private static void AssertCarrierWithinPhysicalPage(PdfReadPage page, PdfTextSpan carrier) {
        var size = page.GetPageSize();
        double endX = carrier.X + carrier.Advance * Math.Cos(carrier.RotationDegrees * Math.PI / 180D);
        double endY = carrier.Y + carrier.Advance * Math.Sin(carrier.RotationDegrees * Math.PI / 180D);
        Assert.InRange(carrier.X, -0.001D, size.Width + 0.001D);
        Assert.InRange(carrier.Y, -0.001D, size.Height + 0.001D);
        Assert.InRange(endX, -0.001D, size.Width + 0.001D);
        Assert.InRange(endY, -0.001D, size.Height + 0.001D);
    }
}
