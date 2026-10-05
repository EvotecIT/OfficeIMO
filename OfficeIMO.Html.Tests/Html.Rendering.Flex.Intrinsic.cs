using OfficeIMO.Html;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
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
