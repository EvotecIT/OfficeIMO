using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("min-content", false)]
    [InlineData("minmax(min-content,40px)", false)]
    [InlineData("fit-content(40px)", false)]
    [InlineData("max-content", true)]
    public void HtmlGrid_NestedFlexKeepsSeparateMinimumAndMaximumContentWidths(string track, bool maximumContent) {
        var options = new HtmlRenderOptions { ViewportWidth = 440D, Margins = HtmlRenderMargins.All(0D) };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf")));
        string html = "<div style='display:grid;width:400px;font:20px Pinned;grid-template-columns:" + track + " 40px;justify-content:start'>"
            + "<div id='nested' style='display:flex;background:red'><div>AAAA BBBB</div></div>"
            + "<div id='tail' style='background:blue'>Tail</div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderShape nested = FindGridShape(rendered, "div#nested");
        HtmlRenderShape tail = FindGridShape(rendered, "div#tail");
        Assert.True(options.Fonts.TryMeasureText("AAAA", 20D, "Pinned", OfficeFontStyle.Regular, out double first));
        Assert.True(options.Fonts.TryMeasureText("BBBB", 20D, "Pinned", OfficeFontStyle.Regular, out double second));
        Assert.True(options.Fonts.TryMeasureText(" ", 20D, "Pinned", OfficeFontStyle.Regular, out double space));
        Assert.Equal(maximumContent ? first + space + second : Math.Max(first, second), nested.Width, 3);
        Assert.Equal(nested.X + nested.Width, tail.X, 3);
    }
}
