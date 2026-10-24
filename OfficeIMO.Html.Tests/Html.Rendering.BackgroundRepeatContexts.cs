using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, "space", "padding-box", "border-box", "right bottom", 2D, 2D, 20.5D, 26D, true)]
    [InlineData(true, "space", "padding-box", "border-box", "right bottom", 2D, 2D, 20.5D, 26D, true)]
    [InlineData(false, "space", "content-box", "border-box", "right bottom", 12D, 18D, 62D / 3D, 14D, false)]
    [InlineData(true, "space", "content-box", "border-box", "right bottom", 12D, 18D, 62D / 3D, 14D, false)]
    [InlineData(false, "space", "border-box", "content-box", "center", 0D, 0D, 21.5D, 15D, true)]
    [InlineData(true, "space", "border-box", "content-box", "center", 0D, 0D, 21.5D, 15D, true)]
    [InlineData(false, "round", "content-box", "border-box", "0% 0%", 12D, 12D, 20D, 20D, true)]
    [InlineData(true, "round", "content-box", "border-box", "0% 0%", 12D, 12D, 20D, 20D, true)]
    [InlineData(false, "round", "content-box", "border-box", "center", 42D, 12D, 20D, 20D, true)]
    [InlineData(true, "round", "content-box", "border-box", "center", 42D, 12D, 20D, 20D, true)]
    [InlineData(false, "round", "content-box", "border-box", "10px 0px", 22D, 12D, 20D, 20D, true)]
    [InlineData(true, "round", "content-box", "border-box", "10px 0px", 22D, 12D, 20D, 20D, true)]
    public void HtmlBackgroundRepeatContexts_UsesPositioningAreaForDistributionAndAuthoredPhase(
        bool margin, string repeat, string origin, string clip, string position,
        double tileX, double tileY, double stepX, double stepY, bool repeatY) {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        string paint = "width:80px;height:20px;padding:10px;border:2px solid green;"
            + "background-image:url(data:image/png;base64," + data + ");background-size:18px 14px;"
            + "background-repeat:" + repeat + ";background-origin:" + origin + ";background-clip:" + clip
            + ";background-position:" + position;
        HtmlRenderDocument rendered = margin
            ? RenderMarginSizing(300D, "@top-center{content:'';" + paint + "}")
            : HtmlRenderTestDriver.Render("<style>html,body{margin:0}</style><div style='" + paint + "'></div>",
                new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderImagePattern pattern = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImagePattern>());
        double clipInset = clip == "content-box" ? 12D : 0D;
        Assert.Equal(pattern.Pattern.Area.X - clipInset + tileX, pattern.Pattern.Tile.X, 6);
        Assert.Equal(pattern.Pattern.Area.Y - clipInset + tileY, pattern.Pattern.Tile.Y, 6);
        Assert.Equal(stepX, pattern.Pattern.HorizontalStep, 6);
        Assert.Equal(stepY, pattern.Pattern.VerticalStep, 6);
        Assert.True(pattern.Pattern.RepeatX);
        Assert.Equal(repeatY, pattern.Pattern.RepeatY);
        rendered.RequireNoLoss();
    }
}
