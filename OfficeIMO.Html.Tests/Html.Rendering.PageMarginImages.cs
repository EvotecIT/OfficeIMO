using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("linear-gradient(to right,red,blue)")]
    [InlineData("repeating-linear-gradient(to right,red 0px 10px,blue 10px 20px)")]
    [InlineData("radial-gradient(ellipse at center,red,blue)")]
    [InlineData("repeating-radial-gradient(circle at center,red 0px 10px,blue 10px 20px)")]
    [InlineData("conic-gradient(from 90deg,red,blue,red)")]
    [InlineData("repeating-conic-gradient(red 0deg 45deg,blue 45deg 90deg)")]
    public void HtmlPageMarginImages_GradientFamiliesPaintTheAllocatedArtifactFrame(string image) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'';background-image:" + image + "}");
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        Assert.Equal(300D, box.Width, 6);
        Assert.Equal(60D, box.Height, 6);
        Assert.Equal(HtmlRenderSemanticGroupRole.Artifact, box.Role);
        Assert.Contains(EnumerateRenderVisuals(box.Visuals), visual => visual is HtmlRenderDrawing
            || visual is HtmlRenderShape shape && (shape.Shape.FillGradient != null || shape.Shape.FillRadialGradient != null));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.NotEqual(OfficeColor.White, raster.GetPixel((int)(box.X + 5), (int)(box.Y + 5)));
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("background-image:linear-gradient(red,blue);background:none")]
    [InlineData("background-image:linear-gradient(red,blue);background-image:initial")]
    [InlineData("background-image:linear-gradient(red,blue);background-image:unset")]
    public void HtmlPageMarginImages_ResetsSuppressPaintInsteadOfInheritingPageImages(string declaration) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "background-image:linear-gradient(green,blue);@top-center{content:'';" + declaration + "}");
        Assert.DoesNotContain(EnumerateRenderVisuals(Assert.Single(rendered.Pages).Scene), visual =>
            visual.Source != null && visual.Source.StartsWith("@page @top-center:background", StringComparison.Ordinal));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("background-image:inherit")]
    [InlineData("background-image:var(--paint)")]
    public void HtmlPageMarginImages_ExplicitInheritanceAndPageVariablesUseTheEffectiveCascade(string declaration) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "--paint:linear-gradient(red,blue);background-image:var(--paint);@top-center{content:'';" + declaration + "}");
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderShape>(), shape => shape.Shape.FillGradient != null);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginImages_LonghandsKeepLayerOrderAndIndependentOriginAndClip() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(12, 8));
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'';width:80px;height:20px;padding:10px;border:2px solid green;"
            + "background-color:red;background-image:url(data:image/png;base64," + data + "),linear-gradient(blue,blue);"
            + "background-position:right bottom;background-size:12px 8px,auto;background-repeat:no-repeat;"
            + "background-origin:content-box;background-clip:padding-box,content-box}");
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        HtmlRenderImage image = Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderImage>());
        Assert.Equal(box.X + 12D + 80D - 12D, image.X, 6);
        Assert.Equal(box.Y + 12D + 20D - 8D, image.Y, 6);
        Assert.Equal(12D, image.Width, 6);
        Assert.Equal(8D, image.Height, 6);
        var paints = EnumerateRenderVisuals(box.Visuals).ToArray();
        Assert.True(Array.FindIndex(paints, visual => visual is HtmlRenderShape shape && shape.Shape.FillGradient != null)
            < Array.FindIndex(paints, visual => visual is HtmlRenderImage));
        Assert.Equal(data, Convert.ToBase64String(image.Bytes));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("content:none;height:20px")]
    [InlineData("content:normal;height:20px")]
    [InlineData("content:'';height:0")]
    [InlineData("content:'';width:0;height:20px")]
    public void HtmlPageMarginImages_NonGeneratingOrZeroAreaBoxesHaveNoPaint(string declaration) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{" + declaration + ";background-image:linear-gradient(red,blue)}");
        Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene), visual =>
            visual.Source != null && visual.Source.StartsWith("@page @top-center:background", StringComparison.Ordinal));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("linear-gradient(red,green,blue)", HtmlRenderDiagnosticCodes.GradientStopLimitExceeded)]
    public void HtmlPageMarginImages_UnsupportedPaintRetainsTheGeneratedSourceAndStrictLoss(string image, string code) {
        string html = "<style>@page{size:300px 300px;margin:60px 20px;@top-center{content:'';background-image:" + image + "}}</style><p>BODY</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { MaxGradientStops = 2 });
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == code && diagnostic.Source == "@page @top-center");
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlPageMarginImages_RepeatExpansionUsesTheOperationBudget() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        string html = "<style>@page{size:300px 300px;margin:60px 20px;@top-center{content:'';"
            + "background-image:url(data:image/png;base64," + data + ");background-size:1px 1px}}</style><p>BODY</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { MaxBackgroundImageTiles = 8 });
        var paints = EnumerateRenderVisuals(PageMarginGeometryBox(rendered.Pages[0], "top-center").Visuals).ToArray();
        Assert.Empty(paints.OfType<HtmlRenderImagePattern>());
        Assert.Single(paints.OfType<HtmlRenderImage>());
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded
            && diagnostic.Source == "@page @top-center");
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlPageMarginImages_CharacterLengthsUseTheSelectedZeroGlyphAdvance() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'';width:80px;height:20px;background-repeat:no-repeat;"
            + "background-image:url(data:image/png;base64," + data + "),linear-gradient(to right,red 0ch,red 2ch,blue 2ch,blue 100%);"
            + "background-size:2ch 10px,auto}", zeroGlyphAdvance: 800);
        HtmlRenderVisual[] paints = EnumerateRenderVisuals(PageMarginGeometryBox(rendered.Pages[0], "top-center").Visuals).ToArray();
        Assert.Equal(32D, Assert.Single(paints.OfType<HtmlRenderImage>()).Width, 6);
        OfficeLinearGradient gradient = Assert.Single(paints.OfType<HtmlRenderShape>(), shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        Assert.Equal(0.4D, gradient.Stops[1].Offset, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginImages_ContainerLengthsRetainTheConfiguredPaperContext() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'';width:80px;height:20px;background-repeat:no-repeat;"
            + "background-image:url(data:image/png;base64," + data + "),linear-gradient(to right,red 0px,red 2cqw,blue 2cqw,blue 100%);"
            + "background-position:1cqw 0;background-size:2cqw 10px,auto}");
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(rendered.Pages[0], "top-center");
        HtmlRenderVisual[] paints = EnumerateRenderVisuals(box.Visuals).ToArray();
        double paperWidth = 210D / 25.4D * 96D;
        HtmlRenderImage image = Assert.Single(paints.OfType<HtmlRenderImage>());
        Assert.Equal(paperWidth * 0.02D, image.Width, 6);
        Assert.Equal(box.X + paperWidth * 0.01D, image.X, 6);
        OfficeLinearGradient gradient = Assert.Single(paints.OfType<HtmlRenderShape>(), shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        Assert.Equal(paperWidth * 0.02D / 80D, gradient.Stops[1].Offset, 6);
        rendered.RequireNoLoss();
    }
}
