using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Theory]
    [InlineData("calc(25% + 1em)", 35D, 70D)]
    [InlineData("calc(10cqw + 1em)", 20D, 40D)]
    [InlineData("calc(var(--portion) + 1em)", 35D, 40D)]
    public void HtmlLengthReuse_RepeatedExpressionsUseEachContainingBlockAndFont(
        string expression, double firstWidth, double secondWidth) {
        string html = "<style>body{margin:0}.sized{width:" + expression
            + ";height:10px;background:red}</style>"
            + "<section style='width:100px;font-size:10px;--portion:25%;container-type:inline-size'><div id='first' class='sized'></div></section>"
            + "<section style='width:200px;font-size:20px;--portion:10cqw;container-type:inline-size'><div id='second' class='sized'></div></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 400D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape[] shapes = rendered.Pages.SelectMany(page => page.Visuals)
            .OfType<HtmlRenderShape>().ToArray();
        Assert.Equal(firstWidth, Assert.Single(shapes, item => item.Source == "div#first"
            && item.Shape.FillColor == OfficeColor.Red).Width, 3);
        Assert.Equal(secondWidth, Assert.Single(shapes, item => item.Source == "div#second"
            && item.Shape.FillColor == OfficeColor.Red).Width, 3);
    }
}
