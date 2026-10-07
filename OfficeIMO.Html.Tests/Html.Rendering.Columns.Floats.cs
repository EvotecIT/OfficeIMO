using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("left", 40D)]
    [InlineData("right", 0D)]
    public void HtmlColumns_FloatExcludesFollowingParagraphsWithoutConsumingTheirFlowHeight(string side, double firstTextX) {
        string html = "<style>body{margin:0;font:12px/20px Arial}p{margin:0}</style>"
            + "<div style='width:220px;height:60px;column-count:2;column-gap:20px;column-fill:auto'>"
            + "<div id='column-float' style='float:" + side + ";width:40px;height:60px;background:red'></div>"
            + "<p>One</p><p>Two</p><p>Three</p><p>Four</p><p>Five</p>"
            + "<p><a href='https://example.test/end'>Six</a></p></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 220D, Margins = HtmlRenderMargins.All(0D), MaxColumnCount = 2
        });
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        Assert.Equal(6, text.Length);
        Assert.Equal(firstTextX, Assert.Single(text, item => item.Text == "One").X, 3);
        Assert.Equal(0D, Assert.Single(text, item => item.Text == "One").Y, 3);
        Assert.Equal(120D, Assert.Single(text, item => item.Text == "Four").X, 3);
        Assert.Equal(0D, Assert.Single(text, item => item.Text == "Four").Y, 3);
        HtmlRenderText last = Assert.Single(text, item => item.Text == "Six");
        Assert.Equal("https://example.test/end", last.LinkUri);
        Assert.Equal(40D, last.Y, 3);
        Assert.All(text, item => Assert.True(item.X + item.Width <= 220D + 0.0001D));
        HtmlRenderShape floating = FindColumnShape(rendered, "div#column-float");
        Assert.Equal(60D, floating.Height, 3);
        Assert.Equal(side == "left" ? 0D : 60D, floating.X, 3);
    }

    [Fact]
    public void HtmlColumns_InlineFloatInParagraphSharesItsExclusionWithFollowingParagraphs() {
        const string html = "<style>body{margin:0;font:12px/20px Arial}p{margin:0}</style>"
            + "<div style='width:220px;height:60px;column-count:2;column-gap:20px;column-fill:auto'>"
            + "<p><span id='nested-column-float' style='float:left;width:40px;height:60px;background:red'></span>One</p>"
            + "<p>Two</p><p>Three</p><p>Four</p><p>Five</p><p>Six</p></div>";
        HtmlRenderDocument rendered = RenderColumns(html, 220D);
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal(6, text.Length);
        Assert.Equal(40D, Assert.Single(text, item => item.Text == "Two").X, 3);
        Assert.Equal(20D, Assert.Single(text, item => item.Text == "Two").Y, 3);
        Assert.Equal(120D, Assert.Single(text, item => item.Text == "Four").X, 3);
        Assert.All(text, item => Assert.True(item.X + item.Width <= 220D + 0.0001D));
        Assert.Equal(60D, FindColumnShape(rendered, "span#nested-column-float").Height, 3);
    }

    [Fact]
    public void HtmlColumns_FloatCrossingColumnBoundaryMovesAsAWholeWithItsAdjacentLines() {
        const string html = "<style>body{margin:0;font:12px/20px Arial}p{margin:0}</style>"
            + "<div style='width:220px;height:60px;column-count:2;column-gap:20px;column-fill:auto'>"
            + "<p>One</p><p>Two</p>"
            + "<div id='deferred-column-float' style='float:left;width:40px;height:60px;background:red'></div>"
            + "<p>Three</p><p>Four</p><p>Five</p></div>";
        HtmlRenderDocument rendered = RenderColumns(html, 220D);
        HtmlRenderShape floating = FindColumnShape(rendered, "div#deferred-column-float");
        Assert.Equal(120D, floating.X, 3);
        Assert.Equal(0D, floating.Y, 3);
        Assert.Equal(60D, floating.Height, 3);
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal(5, text.Length);
        Assert.Equal(160D, Assert.Single(text, item => item.Text == "Three").X, 3);
        Assert.Equal(0D, Assert.Single(text, item => item.Text == "Three").Y, 3);
    }

    [Fact]
    public void HtmlColumns_FloatOnlyPartitionContainsItsPaintBeforeAnAllColumnSpanner() {
        const string html = "<body style='margin:0'><div style='width:220px;column-count:2;column-gap:20px'>"
            + "<div id='only-column-float' style='float:left;width:40px;height:60px;background:red'></div>"
            + "<div id='column-spanner' style='column-span:all;height:20px;background:blue'></div></div>";
        HtmlRenderDocument rendered = RenderColumns(html, 220D);
        Assert.Equal(60D, FindColumnShape(rendered, "div#only-column-float").Height, 3);
        Assert.Equal(60D, FindColumnShape(rendered, "div#column-spanner").Y, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlColumns_NonLineAlignedFloatDoesNotIntroduceATextOrLinkCut(bool paged) {
        const string html = "<style>body{margin:0;font:12px/20px Arial}p{margin:0}</style>"
            + "<div style='width:220px;height:35px;column-count:2;column-gap:20px;column-fill:auto'>"
            + "<div style='float:left;width:40px;height:25px;background:red'></div>"
            + "<p>One</p><p><a href='https://example.test/middle'>Two</a></p><p>Three</p></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = paged ? HtmlRenderMode.Paged : HtmlRenderMode.Continuous,
            ViewportWidth = 220D, PageSize = new OfficePageSize(220D / 96D, 100D / 96D),
            HonorCssPageRules = false, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        Assert.Equal(3, text.Length);
        Assert.Equal("https://example.test/middle", Assert.Single(text, item => item.Text == "Two").LinkUri);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }
}
