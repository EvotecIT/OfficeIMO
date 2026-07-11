using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_MediaLengthsUseTheActiveContinuousSurface() {
        const string html = "<style>"
            + ".target{color:#0000ff}"
            + "@media (max-width:300px){.target{color:#ff0000}}"
            + "@media (min-width:350px) and (max-width:450px){.target{color:#008000}}"
            + "</style><p class='target'>Media marker</p>";

        HtmlRenderDocument medium = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 400D,
            ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderDocument wide = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 800D,
            ViewportHeight = 600D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderText mediumText = Assert.Single(medium.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Media", StringComparison.Ordinal));
        HtmlRenderText wideText = Assert.Single(wide.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Media", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.FromRgb(0, 128, 0), mediumText.Color);
        Assert.Equal(OfficeColor.Blue, wideText.Color);
        Assert.False(HtmlComputedStyleEngine.IsApplicableMedia("(max-width:1px)", HtmlCssMediaContext.Screen, 400D, 200D));
        Assert.True(HtmlComputedStyleEngine.IsApplicableMedia("(orientation:landscape)", HtmlCssMediaContext.Screen, 400D, 200D));
    }

    [Fact]
    public void HtmlRender_MediaLengthsUseTheActivePagedSurface() {
        const string html = "<style>.target{color:#0000ff}@media print and (max-width:300px){.target{color:#ff0000}}</style><p class='target'>Paged media</p>";
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 3D),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text.Contains("Paged", StringComparison.Ordinal));
        Assert.Equal(OfficeColor.Blue, text.Color);
    }
}
