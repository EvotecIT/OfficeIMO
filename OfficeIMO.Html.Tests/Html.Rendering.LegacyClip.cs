using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("absolute", "rect(1px,1px,1px,1px)")]
    [InlineData("fixed", "rect(1px,1px,1px,1px)")]
    [InlineData("absolute", "rect(20px,10px,5px,0)")]
    [InlineData("absolute", "rect(0,5px,10px,20px)")]
    public void HtmlLegacyClip_EmptyRectHidesPaintAndLinksWithoutRemovingSourceDestination(string position, string clip) {
        string html = "<style>*{box-sizing:border-box}body{margin:0}a{position:" + position
            + ";left:10px;top:10px;width:1px;height:1px;padding:8px;border:2px solid red;"
            + "background:red;color:red;overflow:hidden;clip:" + clip + "}</style>"
            + "<a id='skip' href='#main'>Skip</a><div id='main' style='margin-top:60px'>Main</div>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderOptions options = LegacyClipOptions();
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = document.ToPdfBytes(new HtmlToPdfOptions(options));
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.NotNull(document.Document.QuerySelector("#skip"));
        HtmlRenderClipGroup empty = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderClipGroup>(), group => group.Source == "a#skip" && group.ClipWidth == 0D);
        Assert.Equal(0D, empty.ClipHeight);
        Assert.Equal(OfficeColor.White, raster.GetPixel(11, 11));
        Assert.Equal(OfficeColor.White, raster.GetPixel(20, 20));
        Assert.DoesNotContain("html-fragment:main", info.LinkDestinationNames);
        Assert.Contains("html-fragment:main", info.NamedDestinationNames);
        Assert.Empty(rendered.Diagnostics);
    }

    [Theory]
    [InlineData("auto", true)]
    [InlineData("rect(auto,auto,auto,auto)", false)]
    public void HtmlLegacyClip_AutoRectangleClipsOverflowAtBorderBox(string value, bool overflowVisible) {
        string html = "<style>body{margin:0}</style><div style='position:absolute;left:10px;top:10px;"
            + "width:40px;height:30px;box-sizing:border-box;border:2px solid red;padding:3px;clip:" + value + "'>"
            + "<div style='width:60px;height:10px;background:red'></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, LegacyClipOptions());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.Equal(overflowVisible ? OfficeColor.FromRgb(255, 0, 0) : OfficeColor.White, raster.GetPixel(60, 20));
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), raster.GetPixel(11, 20));
        Assert.Empty(rendered.Diagnostics);
    }

    [Theory]
    [InlineData("rect(5px,30px,25px,10px)", 20D, 15D, 20D, 20D)]
    [InlineData("rect(5px 30px 25px 10px)", 20D, 15D, 20D, 20D)]
    [InlineData("rect(auto,auto,auto,10px)", 20D, 10D, 30D, 30D)]
    [InlineData("rect(-5px,auto,auto,-5px)", 5D, 5D, 45D, 35D)]
    [InlineData("rect(calc(1px + 4px),30px,25px,10px)", 20D, 15D, 20D, 20D)]
    public void HtmlLegacyClip_ResolvesOffsetsFromPositionedBorderBox(string value, double x, double y, double width, double height) {
        string html = "<style>body{margin:0}</style><div id='box' style='position:absolute;left:10px;top:10px;"
            + "width:40px;height:30px;box-sizing:border-box;background:red;clip:" + value + "'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, LegacyClipOptions());
        HtmlRenderClipGroup clip = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderClipGroup>(), group => group.Source == "div#box");

        Assert.Equal(x, clip.ClipX, 3);
        Assert.Equal(y, clip.ClipY, 3);
        Assert.Equal(width, clip.ClipWidth, 3);
        Assert.Equal(height, clip.ClipHeight, 3);
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(clip:" + value + ")"));
        Assert.Empty(rendered.Diagnostics);
    }

    [Theory]
    [InlineData("static", "rect(1px,1px,1px,1px)")]
    [InlineData("relative", "rect(1px,1px,1px,1px)")]
    [InlineData("static", "rect(10%,20%,30%,40%)")]
    [InlineData("absolute", "auto")]
    public void HtmlLegacyClip_AutoAndInapplicablePositionsKeepPaint(string position, string clip) {
        string html = "<style>body{margin:0}</style><div style='position:" + position
            + ";width:40px;height:30px;background:red;clip:" + clip + "'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, LegacyClipOptions());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), raster.GetPixel(10, 10));
        Assert.Empty(rendered.Diagnostics);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlLegacyClip_IntersectsClipPathInsideTransformedPaintAndLinkArea(bool linkedSvg) {
        const string uri = "https://example.test/clip";
        string html = "<style>body{margin:0}</style><a href='" + uri + "' style='position:absolute;left:10px;top:10px;"
            + "display:block;width:40px;height:30px;background:red;clip:rect(0,30px,30px,10px);"
            + "clip-path:inset(10px 0 0);transform:translateX(10px)'>";
        if (linkedSvg) html += "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='30' style='display:block'>"
            + "<rect width='40' height='30' fill='red'/></svg>";
        html += "</a>";
        HtmlRenderOptions options = LegacyClipOptions();
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(options));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(uri));

        Assert.Equal(OfficeColor.White, raster.GetPixel(25, 25));
        Assert.Equal(OfficeColor.White, raster.GetPixel(35, 15));
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), raster.GetPixel(35, 25));
        Assert.Equal(OfficeColor.White, raster.GetPixel(55, 25));
        Assert.InRange(link.X1, 22.49D, 22.51D);
        Assert.InRange(link.Width, 14.99D, 15.01D);
        Assert.InRange(link.Height, 14.99D, 15.01D);
        Assert.Empty(rendered.Diagnostics);
    }

    [Fact]
    public void HtmlPdf_RectangularClipPathRetainsPaintedBackgroundOnlyAnchor() {
        const string uri = "https://example.test/background-anchor";
        const string html = "<style>body{margin:0}</style><a href='" + uri
            + "' style='display:block;width:40px;height:30px;background:red;clip-path:inset(0)'></a>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(LegacyClipOptions()));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(uri));

        Assert.InRange(link.X1, 0D, .01D);
        Assert.InRange(link.Width, 29.99D, 30.01D);
        Assert.InRange(link.Height, 22.49D, 22.51D);
    }

    [Theory]
    [InlineData("clip:rect(1px,1px,1px,1px)", 1)]
    [InlineData("clip-path:inset(100%)", 1)]
    [InlineData("clip-path:inset(90px 0 0 90px)", 2)]
    public void HtmlPdf_ClippedAnchorDoesNotSuppressVisibleEqualUriSibling(string clipping, int annotations) {
        const string uri = "https://example.test/shared";
        string html = "<style>body{margin:0}</style><a href='" + uri
            + "' style='position:absolute;left:0;top:0;display:block;width:100px;height:100px;background:red;"
            + clipping + "'></a><a href='" + uri
            + "' style='display:block;width:80px;height:30px;background:green'>Visible</a>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(LegacyClipOptions()));
        IReadOnlyList<PdfCore.PdfLogicalLinkAnnotation> links = PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri(uri);

        Assert.Equal(annotations, links.Count);
        Assert.Contains(links, link => link.X1 <= .01D && link.Width >= 59.99D && link.Height >= 22.49D);
    }

    [Theory]
    [InlineData("rect(0,20px,10px,0)")]
    [InlineData("rect(auto,auto,auto,auto)")]
    [InlineData("rect(1px,1px,1px,1px)")]
    public void HtmlLegacyClip_ControlsUseDiagnosedStaticPaint(string value) {
        string html = "<div id='clip' style='position:absolute;left:10px;top:10px;width:80px;height:40px;clip:" + value
            + "'><input id='child' name='child' value='Field' style='width:60px;height:20px'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, LegacyClipOptions());
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics,
            item => item.Code == HtmlRenderDiagnosticCodes.FormFieldTransformStaticFallback);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(LegacyClipOptions()));

        Assert.Equal("input#child", diagnostic.Source);
        Assert.StartsWith("ancestor-clip=", diagnostic.Detail);
        Assert.True(rendered.HasLoss);
        Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene), visual => visual is HtmlRenderFormField);
        Assert.Empty(PdfCore.PdfDocumentReadResult.Load(pdf).FormFields);
        HtmlRenderOptions strict = LegacyClipOptions();
        strict.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
        Assert.Throws<HtmlConversionException>(() => HtmlRenderTestDriver.Render(html, strict));
    }

    [Theory]
    [InlineData("rect(0,50%,20px,0)")]
    [InlineData("rect(0,30px 20px,0)")]
    [InlineData("circle(5px)")]
    public void HtmlLegacyClip_UnsupportedApplicableValuesProduceTypedLoss(string value) {
        string html = "<div id='unsupported' style='position:absolute;width:40px;height:30px;clip:" + value + "'>Visible</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, LegacyClipOptions());
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics,
            item => item.Code == HtmlRenderDiagnosticCodes.ClipValueUnsupported);

        Assert.Equal("div#unsupported", diagnostic.Source);
        Assert.True(rendered.HasLoss);
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(clip:" + value + ")"));
        HtmlRenderOptions strict = LegacyClipOptions();
        strict.FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss;
        Assert.Throws<HtmlConversionException>(() => HtmlRenderTestDriver.Render(html, strict));
    }

    private static HtmlRenderOptions LegacyClipOptions() => new() {
        ViewportWidth = 120D,
        ViewportHeight = 100D,
        Margins = HtmlRenderMargins.All(0D)
    };
}
