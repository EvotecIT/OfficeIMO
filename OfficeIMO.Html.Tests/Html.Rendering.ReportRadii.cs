using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("19.9999999px", "10px")]
    [InlineData("19.99996px", "10px / 9.99998px")]
    [InlineData("20px", "100px")]
    public void HtmlReportRadii_FractionalBoxesPreserveRoundedFillAndClip(string height, string radius) {
        string html = "<div id='rounded' style='width:100px;height:" + height
            + ";border-radius:" + radius + ";background:red;overflow:hidden'>Clipped content</div>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 120D, ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D), BackgroundColor = OfficeColor.Transparent
        });
        HtmlRenderShape fill = Assert.Single(document.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#rounded");

        Assert.Equal(OfficeShapeKind.RoundedRectangle, fill.Shape.Kind);
        Assert.InRange(fill.Shape.CornerRadius, 0.1D, Math.Min(fill.Shape.Width, fill.Shape.Height) / 2D);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(document.Pages[0].CreateDrawing());
        Assert.True(image.Width > 0 && image.Height > 0);
        Assert.DoesNotContain(document.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported);
    }

    [Fact]
    public void HtmlReportRadii_InsetClipAcceptsNearlyCircularAxes() {
        const string html = "<div style='width:100px;height:19.99996px;clip-path:inset(0 round 10px / 9.99998px);background:red'>Clipped</div>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 120D, ViewportHeight = 40D, Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains(document.Pages[0].Visuals, visual => visual is HtmlRenderPathClipGroup);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(document.Pages[0].CreateDrawing());
        Assert.True(image.Width > 0 && image.Height > 0);
    }

    [Theory]
    [InlineData("<input type='range' style='width:1px;height:20px'>")]
    [InlineData("<progress value='0.5' style='width:1px;height:20px'></progress>")]
    [InlineData("<meter value='0.5' style='width:1px;height:20px'></meter>")]
    public void HtmlReportRadii_NarrowControlTracksFitBothDimensions(string html) {
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 120D, ViewportHeight = 40D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape track = Assert.Single(document.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            visual => visual.Source.EndsWith(":track", StringComparison.Ordinal));

        Assert.Equal(OfficeShapeKind.RoundedRectangle, track.Shape.Kind);
        Assert.InRange(track.Shape.CornerRadius, 0D, Math.Min(track.Shape.Width, track.Shape.Height) / 2D);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(document.Pages[0].CreateDrawing());
        Assert.True(image.Width > 0 && image.Height > 0);
    }
}
