using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("absolute")]
    [InlineData("fixed")]
    public void HtmlLegacyClip_EmptyAncestorClipsExtractedFixedPaintAndControls(string position) {
        string html = "<style>body{margin:0}</style><div style='position:" + position
            + ";left:10px;top:10px;width:100px;height:100px;clip:rect(0,0,0,0)'>"
            + FixedClipLink() + "<input id='field' name='field' value='Field' style='position:fixed;left:20px;top:70px;width:60px;height:20px'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, LegacyClipOptions());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(LegacyClipOptions()));

        Assert.Equal(OfficeColor.White, raster.GetPixel(30, 30));
        Assert.Equal(OfficeColor.White, raster.GetPixel(30, 80));
        Assert.Empty(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(FixedClipUri));
        Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldTransformStaticFallback);
    }

    [Fact]
    public void HtmlLegacyClip_FixedPaintIntersectsAncestorBorderCoordinates() {
        string html = "<style>body{margin:0}</style><div style='position:absolute;left:10px;top:10px;"
            + "width:100px;height:100px;clip:rect(10px,50px,40px,20px)'>" + FixedClipLink() + "</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, LegacyClipOptions());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(LegacyClipOptions()));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(FixedClipUri));

        Assert.Equal(OfficeColor.White, raster.GetPixel(25, 25));
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), raster.GetPixel(40, 30));
        Assert.Equal(OfficeColor.White, raster.GetPixel(65, 30));
        Assert.Equal(22.5D, link.X1, 6);
        Assert.Equal(22.5D, link.Width, 6);
        Assert.Equal(22.5D, link.Height, 6);
    }

    [Theory]
    [InlineData("absolute")]
    [InlineData("fixed")]
    public void HtmlLegacyClip_NestedPositionedAncestorsIntersectFixedPaint(string innerPosition) {
        string html = "<style>body{margin:0}</style><div style='position:absolute;left:10px;top:10px;"
            + "width:100px;height:100px;clip:rect(0,60px,60px,0)'><div style='position:" + innerPosition
            + ";left:" + (innerPosition == "fixed" ? "20px;top:20px;" : "10px;top:10px;")
            + "width:80px;height:80px;clip:rect(0,20px,20px,0)'>" + FixedClipLink() + "</div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, LegacyClipOptions());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(LegacyClipOptions()));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(FixedClipUri));

        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), raster.GetPixel(30, 30));
        Assert.Equal(OfficeColor.White, raster.GetPixel(45, 30));
        Assert.Equal(15D, link.X1, 6);
        Assert.Equal(15D, link.Width, 6);
        Assert.Equal(15D, link.Height, 6);
    }

    [Fact]
    public void HtmlLegacyClip_AncestorIntersectionRetainsFixedChildTransformAndClipPath() {
        string html = "<style>body{margin:0}</style><div style='position:absolute;left:10px;top:10px;"
            + "width:100px;height:100px;clip:rect(10px,50px,40px,20px);clip-path:inset(20px 0 0)'>"
            + FixedClipLink("transform:translateX(15px)") + "</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, LegacyClipOptions());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(LegacyClipOptions()));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(FixedClipUri));

        Assert.Equal(OfficeColor.White, raster.GetPixel(25, 35));
        Assert.Equal(OfficeColor.White, raster.GetPixel(40, 25));
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), raster.GetPixel(40, 35));
        Assert.Equal(26.25D, link.X1, 6);
        Assert.Equal(18.75D, link.Width, 6);
        Assert.Equal(15D, link.Height, 6);
    }

    [Theory]
    [InlineData("translateX(15px)", true)]
    [InlineData("scale(0)", false)]
    public void HtmlLegacyClip_TransformedAncestorClipConstrainsLocallyFixedPaint(string transform, bool visible) {
        string html = "<style>body{margin:0}</style><div style='position:absolute;left:10px;top:10px;"
            + "width:100px;height:100px;clip:rect(10px,50px,40px,20px);transform:" + transform
            + ";transform-origin:0 0'>" + FixedClipLink() + "</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, LegacyClipOptions());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(LegacyClipOptions()));
        var links = PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(FixedClipUri);

        Assert.Equal(OfficeColor.White, raster.GetPixel(40, 30));
        Assert.Equal(visible ? OfficeColor.FromRgb(255, 0, 0) : OfficeColor.White, raster.GetPixel(50, 30));
        if (visible) {
            PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(links);
            Assert.Equal(33.75D, link.X1, 6);
            Assert.Equal(22.5D, link.Width, 6);
            Assert.Equal(15D, link.Height, 6);
        } else Assert.Empty(links);
    }

    [Fact]
    public void HtmlLegacyClip_OverflowOnlyAncestorDoesNotClipViewportFixedPaint() {
        string html = "<style>body{margin:0}</style><div style='position:absolute;left:10px;top:10px;"
            + "width:1px;height:1px;overflow:hidden'>" + FixedClipLink() + "</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, LegacyClipOptions());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(LegacyClipOptions()));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(FixedClipUri));

        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), raster.GetPixel(30, 30));
        Assert.Equal(45D, link.Width, 6);
        Assert.Equal(30D, link.Height, 6);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, true)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, true)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, true)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, false)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, false)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, false)]
    public void HtmlLegacyClip_FixedClippingSurvivesPagedAndSnapshotProjection(HtmlRenderIntentProfile profile, bool empty) {
        string html = "<style>html,body{margin:0;height:250px}</style><div style='height:250px'></div>"
            + "<div style='position:absolute;left:10px;top:10px;width:100px;height:100px;clip:"
            + (empty ? "rect(0,0,0,0)" : "rect(10px,50px,40px,20px)") + "'>" + FixedClipLink() + "</div>";
        var options = new HtmlToPdfOptions { PageSize = new OfficePageSize(120D / 96D, 100D / 96D),
            ViewportWidth = 120, ViewportHeight = 250, Margins = HtmlRenderMargins.All(0), AutoFitWidePrintContent = false };
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderRequest request = HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, options);
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(source, request).Document;
        byte[] pdf = source.RenderToPdfBytes(request);
        var links = PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(FixedClipUri);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.Equal(empty ? 0 : profile == HtmlRenderIntentProfile.ScreenSnapshotPaged ? 1 : 3, links.Count);
        for (int index = 0; index < rendered.Pages.Count; index++) {
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[index].CreateDrawing());
            bool visible = !empty && (index == 0 || profile != HtmlRenderIntentProfile.ScreenSnapshotPaged);
            Assert.Equal(visible ? OfficeColor.FromRgb(255, 0, 0) : OfficeColor.White, raster.GetPixel(40, 30));
            Assert.Equal(OfficeColor.White, raster.GetPixel(25, 25));
        }
    }

    private const string FixedClipUri = "https://example.test/fixed";
    private static string FixedClipLink(string extra = "") => "<a href='" + FixedClipUri
        + "' style='position:fixed;left:20px;top:20px;width:60px;height:40px;background:red;" + extra + "'></a>";
}
