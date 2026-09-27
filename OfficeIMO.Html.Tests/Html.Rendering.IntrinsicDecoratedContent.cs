using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("inline-block", "svg")]
    [InlineData("inline-block", "generated")]
    [InlineData("positioned", "svg")]
    [InlineData("positioned", "generated")]
    [InlineData("float", "svg")]
    [InlineData("float", "generated")]
    public void HtmlRendering_ShrinkToFitIncludesDecoratedInlineContent(string mode, string decoration) {
        string placement = mode switch {
            "positioned" => "position:absolute;left:0;top:0;",
            "float" => "float:left;",
            _ => "display:inline-block;"
        };
        string icon = decoration == "svg"
            ? "<svg width='24' height='24' viewBox='0 0 24 24'><circle cx='12' cy='12' r='8'/></svg>"
            : "<span class='icon'></span>";
        string html = "<style>body{margin:0}.badge{white-space:nowrap;padding:4px;background:blue}"
            + ".icon::before{content:'WW';font-size:24px}</style>"
            + "<div style='position:relative;height:60px'><span id='plain' class='badge' style='" + placement + "'>Save</span></div>"
            + "<div style='position:relative;height:60px'><span id='mixed' class='badge' style='" + placement + "'>" + icon + "Save</span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 400D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape[] shapes = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Visuals))
            .OfType<HtmlRenderShape>().Where(item => item.Shape.FillColor.HasValue).ToArray();
        HtmlRenderShape plain = Assert.Single(shapes, item => item.Source == "span#plain");
        HtmlRenderShape mixed = Assert.Single(shapes, item => item.Source == "span#mixed");
        Assert.True(mixed.Width >= plain.Width + 23D,
            $"The decorated badge width {mixed.Width} must include its icon in addition to the plain label width {plain.Width}.");
    }
}
