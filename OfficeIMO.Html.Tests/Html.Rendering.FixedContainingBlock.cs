using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("position:absolute;left:100px;top:50px", false)]
    [InlineData("position:absolute;left:100px;top:50px", true)]
    [InlineData("position:relative;margin-left:100px;margin-top:50px", false)]
    [InlineData("position:relative;margin-left:100px;margin-top:50px", true)]
    [InlineData("margin-left:100px;margin-top:50px", false)]
    public void HtmlFixedPosition_TransformedAncestorUsesLocalPaddingBox(string placement, bool clip) {
        string html = FixedContainingHeader + "<div style='" + placement
            + ";width:100px;height:80px;transform:translateX(15px);transform-origin:0 0;"
            + (clip ? "clip:rect(0,100px,80px,0)" : "") + "'>"
            + FixedContainingLink("left:20px;top:10px;width:40px;height:20px") + "</div>";

        AssertFixedContainingPaintAndPdf(html, 135, 60, 40, 20);
    }

    [Theory]
    [InlineData("fixed")]
    [InlineData("absolute")]
    public void HtmlPositioning_TransformedStaticAncestorResolvesPercentInsetsAgainstPaddingBox(string position) {
        string html = FixedContainingHeader + "<div style='margin-left:100px;margin-top:50px;"
            + "width:100px;height:80px;padding:10px;border:5px solid transparent;"
            + "transform:translateX(15px);transform-origin:0 0'><a href='" + FixedContainingUri
            + "' style='position:" + position + ";right:10%;bottom:10%;width:25%;height:20%;background:red'></a></div>";

        AssertFixedContainingPaintAndPdf(html, 198, 125, 30, 20);
    }

    [Fact]
    public void HtmlFixedPosition_NearestNestedTransformRetainsChildTransformAndAuthoredClip() {
        string html = FixedContainingHeader + "<div style='position:absolute;left:100px;top:50px;"
            + "width:100px;height:80px;clip:rect(0,100px,80px,0);transform:translateX(15px);transform-origin:0 0'>"
            + "<div style='margin-left:10px;margin-top:5px;width:80px;height:60px;transform:translate(7px,3px);transform-origin:0 0'>"
            + FixedContainingLink("left:20px;top:10px;width:40px;height:20px;transform:translate(4px,2px);"
                + "transform-origin:0 0;clip:rect(0,30px,20px,0);clip-path:inset(0)") + "</div></div>";

        AssertFixedContainingPaintAndPdf(html, 156, 70, 30, 20);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlLegacyClip_IntermediatePositionedAncestorClipsLocallyFixedGrandchild(bool empty) {
        string html = FixedContainingHeader + "<div style='position:absolute;left:100px;top:50px;"
            + "width:100px;height:80px;transform:translateX(15px);transform-origin:0 0'>"
            + "<div style='position:absolute;left:10px;top:5px;width:80px;height:60px;clip:"
            + (empty ? "rect(0,0,0,0)" : "rect(5px,40px,25px,10px)") + "'>"
            + FixedContainingLink("left:20px;top:10px;width:40px;height:20px") + "</div></div>";
        if (!empty) {
            AssertFixedContainingPaintAndPdf(html, 135, 60, 30, 20);
            return;
        }
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedContainingOptions());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(FixedContainingOptions()));

        Assert.Equal(OfficeColor.White, raster.GetPixel(140, 65));
        Assert.Empty(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(FixedContainingUri));
    }

    [Fact]
    public void HtmlFixedPosition_TransformedGridUsesAuthoredGridArea() {
        string html = FixedContainingHeader + "<div style='margin-left:100px;margin-top:50px;display:grid;"
            + "grid-template-columns:40px 60px;grid-template-rows:80px;width:100px;height:80px;"
            + "padding:10px;border:5px solid transparent;transform:translateX(15px);transform-origin:0 0'>"
            + FixedContainingLink("grid-column:2 / 3;grid-row:1 / 2;left:20px;top:10px;width:40px;height:20px") + "</div>";

        AssertFixedContainingPaintAndPdf(html, 190, 75, 40, 20);
    }

    [Fact]
    public void HtmlFixedPosition_TransformedBodyCreatesLocalContainingBox() {
        string html = "<style>html,body{margin:0}body{width:200px;height:100px;transform:translateX(15px);transform-origin:0 0}</style>"
            + FixedContainingLink("left:20px;top:10px;width:40px;height:20px");

        AssertFixedContainingPaintAndPdf(html, 35, 10, 40, 20);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    public void HtmlFixedPosition_TransformedLocalBoxContinuesWithoutViewportRepetition(HtmlRenderIntentProfile profile) {
        string html = "<style>html,body{margin:0}</style><div style='width:100px;height:250px;transform:translateX(15px);transform-origin:0 0'>"
            + FixedContainingLink("left:20px;top:120px;width:40px;height:20px") + "<div style='height:250px'></div></div>"
            + "<a href='https://example.test/viewport-fixed' style='position:fixed;left:80px;top:10px;width:10px;height:10px;background:blue'></a>";
        var options = new HtmlToPdfOptions { PageSize = new OfficePageSize(120D / 96D, 100D / 96D),
            ViewportWidth = 120, ViewportHeight = 250, Margins = HtmlRenderMargins.All(0), AutoFitWidePrintContent = false };
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderRequest request = HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, options);
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(source, request).Document;
        PdfCore.PdfDocumentReadResult pdf = PdfCore.PdfDocumentReadResult.Load(source.RenderToPdfBytes(request));

        Assert.Equal(3, rendered.Pages.Count);
        Assert.Single(pdf.GetLinksByUri(FixedContainingUri));
        Assert.Equal(profile == HtmlRenderIntentProfile.ScreenSnapshotPaged ? 1 : 3,
            pdf.GetLinksByUri("https://example.test/viewport-fixed").Count);
        for (int index = 0; index < rendered.Pages.Count; index++) {
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[index].CreateDrawing());
            Assert.Equal(index == 1 ? OfficeColor.FromRgb(255, 0, 0) : OfficeColor.White, raster.GetPixel(40, 25));
            Assert.Equal(index == 0 || profile != HtmlRenderIntentProfile.ScreenSnapshotPaged
                ? OfficeColor.FromRgb(0, 0, 255) : OfficeColor.White, raster.GetPixel(85, 15));
        }
    }

    private static void AssertFixedContainingPaintAndPdf(string html, double x, double y, double width, double height) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedContainingOptions());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(FixedContainingOptions()));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(FixedContainingUri));

        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), raster.GetPixel((int)x + 5, (int)y + 5));
        Assert.Equal(OfficeColor.White, raster.GetPixel((int)(x + width) + 2, (int)y + 5));
        Assert.Equal(x * .75D, link.X1, 6);
        Assert.Equal(width * .75D, link.Width, 6);
        Assert.Equal(height * .75D, link.Height, 6);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PositionStaticAnchorFallback);
    }

    private const string FixedContainingHeader = "<style>html,body{margin:0}</style>";
    private const string FixedContainingUri = "https://example.test/local-fixed";
    private static HtmlRenderOptions FixedContainingOptions() => new() {
        ViewportWidth = 450, ViewportHeight = 200, Margins = HtmlRenderMargins.All(0) };
    private static string FixedContainingLink(string style) => "<a href='" + FixedContainingUri
        + "' style='position:fixed;background:red;" + style + "'></a>";
}
