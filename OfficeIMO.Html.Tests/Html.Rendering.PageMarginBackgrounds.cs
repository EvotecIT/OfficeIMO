using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("background-color:color(from blue srgb r g b)", "background-color")]
    [InlineData("--paint:color(from blue srgb r g b);background-color:var(--paint)", "background-color")]
    [InlineData("color:color(from blue srgb r g b);background-color:currentcolor", "color")]
    public void HtmlPageMarginBackgrounds_UnsupportedColorsRetainTheirTypedLossContract(string declaration, string property) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D, "@top-center{content:'A';" + declaration + "}");
        var diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ColorValueUnsupported);
        Assert.Equal("@page @top-center", diagnostic.Source);
        Assert.Equal(property + "=color(from blue srgb r g b)", diagnostic.Detail);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("background-color:red")]
    [InlineData("background:red")]
    [InlineData("background:blue;background-color:red")]
    [InlineData("background-color:blue;background:red")]
    [InlineData("background-color:red!important;background:blue")]
    [InlineData("--paint:red;background:var(--paint)")]
    [InlineData("background-color:red;background-color:invalid-color")]
    public void HtmlPageMarginBackgrounds_CascadePaintsTheWholeAllocatedFrame(string declaration) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'A';" + declaration + "}@bottom-center{content:'B';" + declaration + "}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        foreach (string position in new[] { "top-center", "bottom-center" }) {
            HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, position);
            HtmlRenderShape fill = Assert.Single(box.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.FillColor == OfficeColor.Red);
            Assert.Equal(300D, fill.Width, 6);
            Assert.Equal(60D, fill.Height, 6);
            Assert.Equal(box.X, fill.X, 6);
            Assert.Equal(box.Y, fill.Y, 6);
            Assert.Equal(HtmlRenderSemanticGroupRole.Artifact, box.Role);
            Assert.IsType<HtmlRenderShape>(box.Visuals.First());
            Assert.Contains(box.Visuals, visual => visual is HtmlRenderText);
        }
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("background:blue;background:none")]
    [InlineData("background-color:red;background:initial")]
    [InlineData("background-color:red;background-color:unset")]
    [InlineData("background-color:transparent")]
    [InlineData("")]
    public void HtmlPageMarginBackgrounds_TransparentDefaultsAndResetsDoNotInheritPagePaint(string declaration) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "background-color:blue;@top-center{content:'A';" + declaration + "}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, "top-center");
        Assert.Empty(box.Visuals.OfType<HtmlRenderShape>());
        Assert.Contains(page.Scene.OfType<HtmlRenderShape>(), shape => shape.Shape.FillColor == OfficeColor.Blue);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("color:red;background-color:currentcolor", false)]
    [InlineData("background-color:inherit", true)]
    public void HtmlPageMarginBackgrounds_CurrentColorAndExplicitPageInheritanceUseTheGeneratedContext(string declaration, bool inherits) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "background-color:blue;color:green;@top-center{content:'';height:20px;" + declaration + "}");
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        HtmlRenderShape fill = Assert.Single(box.Visuals.OfType<HtmlRenderShape>());
        Assert.Equal(inherits ? OfficeColor.Blue : OfficeColor.Red, fill.Shape.FillColor);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("border-box", 0D, 0D, 104D, 44D)]
    [InlineData("padding-box", 2D, 2D, 100D, 40D)]
    [InlineData("content-box", 12D, 12D, 80D, 20D)]
    [InlineData("CONTENT-BOX", 12D, 12D, 80D, 20D)]
    [InlineData("padding-box,content-box", 2D, 2D, 100D, 40D)]
    public void HtmlPageMarginBackgrounds_SolidClipReusesTheBorderPaddingAndContentFrames(
        string clip, double dx, double dy, double width, double height) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'';width:80px;height:20px;padding:10px;border:2px solid green;background:red;background-clip:" + clip + "}");
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        HtmlRenderShape fill = Assert.Single(box.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.FillColor == OfficeColor.Red);
        Assert.Equal(box.X + dx, fill.X, 6);
        Assert.Equal(box.Y + dy, fill.Y, 6);
        Assert.Equal(width, fill.Width, 6);
        Assert.Equal(height, fill.Height, 6);
        Assert.IsType<HtmlRenderShape>(box.Visuals.First());
        Assert.Contains(box.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.StrokeColor == OfficeColor.Green);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("none")]
    [InlineData("normal")]
    public void HtmlPageMarginBackgrounds_NonGeneratingContentNeverPaints(string content) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:" + content + ";height:20px;background:red}");
        Assert.DoesNotContain(Assert.Single(rendered.Pages).Scene.OfType<HtmlRenderSemanticGroup>(), group => group.Source == "@page @top-center");
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("width:0;height:20px")]
    [InlineData("width:80px;height:0")]
    public void HtmlPageMarginBackgrounds_ZeroAreaHasNoSolidPaint(string dimensions) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'';background:red;" + dimensions + "}");
        Assert.Empty(PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center").Visuals);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginBackgrounds_ZeroFontRetainsSolidAndGradientPaint() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'TEXT';font-size:0;height:20px;background:linear-gradient(blue,green) red}");
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        Assert.Equal(2, EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderShape>().Count());
        Assert.Empty(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>());
        rendered.RequireNoLoss();
    }
}
