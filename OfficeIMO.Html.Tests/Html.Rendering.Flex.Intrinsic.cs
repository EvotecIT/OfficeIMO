using OfficeIMO.Html;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("row")]
    [InlineData("column")]
    public void HtmlFlexIntrinsic_PreservesTheOriginatingFontOfCollapsedWhitespace(string direction) {
        var options = new HtmlRenderOptions { ViewportWidth = 640D, Margins = HtmlRenderMargins.All(0D) };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf")));
        string html = "<div style='display:flex;flex-direction:" + direction + ";align-items:flex-start;width:600px;font:40px Pinned'>"
            + "<span id='item' style='background:#eeeeee'>AAAA <span style='font-size:10px'>BBBB</span></span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText large = Assert.Single(text, item => item.Text.Trim() == "AAAA");
        HtmlRenderText small = Assert.Single(text, item => item.Text.Trim() == "BBBB");
        Assert.True(small.Y + small.Height <= large.Y + large.Height + 0.001D);
    }

    [Theory]
    [InlineData("row")]
    [InlineData("column")]
    public void HtmlFlexIntrinsic_PreservesNestedRowWidthsAndGaps(string direction) {
        var options = new HtmlRenderOptions { ViewportWidth = 640D, Margins = HtmlRenderMargins.All(0D) };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf")));
        string html = "<div style='display:flex;flex-direction:" + direction + ";align-items:flex-start;width:600px;font:20px Pinned'>"
            + "<div id='inner' style='display:flex;flex-shrink:0;gap:8px;background:#eeeeee'>"
            + "<div>AAAA</div><div>BBBB</div></div><span>Tail</span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape inner = FindFlexShape(rendered, "div#inner");
        Assert.True(options.Fonts.TryMeasureText("AAAA", 20D, "Pinned", OfficeFontStyle.Regular, out double first));
        Assert.True(options.Fonts.TryMeasureText("BBBB", 20D, "Pinned", OfficeFontStyle.Regular, out double second));
        Assert.Equal(first + second + 8D, inner.Width, 3);
        if (direction == "row") {
            HtmlRenderText tail = Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), item => item.Text == "Tail");
            Assert.True(tail.X >= inner.X + inner.Width - 0.001D);
        }
    }

    [Theory]
    [InlineData("row")]
    [InlineData("column")]
    public void HtmlFlexIntrinsic_KeepsKernedWordsAtTheirNaturalWidth(string direction) {
        var options = new HtmlRenderOptions { ViewportWidth = 640D, Margins = HtmlRenderMargins.All(0D) };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf")));
        string html = "<div style='display:flex;flex-direction:" + direction
            + ";align-items:flex-start;width:600px;font:12px Pinned'>"
            + "<span id='item' style='background:#eeeeee'>Word Word</span></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        Assert.NotEmpty(text);
        Assert.Single(text.Select(item => Math.Round(item.Y, 3)).Distinct());
        Assert.Equal("Word Word", string.Concat(text.Select(item => item.Text)).Trim());
    }

    [Theory]
    [InlineData("row")]
    [InlineData("column")]
    public void HtmlFlexIntrinsic_KeepsGeneratedContentOnItsNaturalLine(string direction) {
        string html = "<style>#item::before{content:'Prefix '}#item::after{content:' Suffix'}</style>"
            + "<div style='display:flex;flex-direction:" + direction
            + ";align-items:flex-start;flex-wrap:wrap;width:600px;font-size:12px'>"
            + "<span id='item' style='background:#eeeeee'>Body</span><span>Next</span></div>";

        HtmlRenderDocument rendered = RenderFlex(html, 640D);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText before = Assert.Single(text, item => item.Source == "span#item::before");
        HtmlRenderText body = Assert.Single(text, item => item.Text == "Body");
        HtmlRenderText after = Assert.Single(text, item => item.Source == "span#item::after");

        Assert.Equal("Prefix", before.Text.Trim());
        Assert.Equal("Suffix", after.Text.Trim());
        Assert.Equal(before.Y, body.Y, 3);
        Assert.Equal(body.Y, after.Y, 3);
        Assert.True(before.X < body.X);
        Assert.True(body.X < after.X);
    }
}
