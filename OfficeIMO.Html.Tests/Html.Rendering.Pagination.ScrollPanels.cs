using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("")]
    [InlineData("max-height:100vh;")]
    [InlineData("height:300px;")]
    public void HtmlRender_PagedScrollPanelUsesRemainingPageSpace(string sizing) {
        HtmlRenderDocument rendered = RenderScrollPanel(sizing);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Text == "Panel item 1");
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>(),
            text => text.Text == "Panel item 6");
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        for (int index = 1; index <= 6; index++)
            Assert.Single(text, item => item.Text == $"Panel item {index}");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
    }

    [Fact]
    public void HtmlRender_PagedScrollPanelHonorsAnAuthoredKeepTogether() {
        HtmlRenderDocument rendered = RenderScrollPanel("max-height:100vh;break-inside:avoid;");

        Assert.Equal(2, rendered.Pages.Count);
        Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Text.StartsWith("Panel item", StringComparison.Ordinal));
        Assert.Equal(6, EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderText>()
            .Count(text => text.Text.StartsWith("Panel item", StringComparison.Ordinal)));
    }

    private static HtmlRenderDocument RenderScrollPanel(string sizing) {
        string rows = string.Concat(Enumerable.Range(1, 6).Select(index => $"<p>Panel item {index}</p>"));
        string html = "<style>@page{size:400px 500px;margin:0}html,body{margin:0;font:14px Arial;line-height:20px}"
            + "p{margin:0;padding-bottom:30px}</style><div style='height:250px'>Report title</div>"
            + "<section style='overflow-y:auto;" + sizing + "'>" + rows + "</section>";
        return HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true, Margins = HtmlRenderMargins.All(0D)
        });
    }
}
