using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void PagedRendererFragmentsFittingFooterWhenEarlierPageHasRoom() {
        const string html = "<style>@page{size:300px 200px;margin:0}html,body{margin:0}"
            + "#lead{height:120px}footer div{height:50px;font:12px Arial}</style>"
            + "<div id='lead'>Lead</div><footer><div>Feedback</div>"
            + "<div>Social links</div><div>Agency links</div></footer>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, HonorCssPageRules = true });

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Feedback");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Agency links");
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(20, 2)]
    public void PagedRendererKeepsFinalContentWithBottomPadding(int bottomPadding, int expectedPages) {
        string html = "<style>@page{size:100px 100px;margin:0}html,body{margin:0;padding:0}"
            + "#lead{height:70px}#footer{font:10px/10px Arial;background:#ddd;padding:5px 0 "
            + bottomPadding + "px}</style><div id='lead'>Lead</div><div id='footer'>Footer</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, HonorCssPageRules = true });

        Assert.Equal(expectedPages, rendered.Pages.Count);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.Text.Contains("Footer", StringComparison.Ordinal) && bottomPadding > 0);
        Assert.Contains(rendered.Pages[rendered.Pages.Count - 1].Visuals.OfType<HtmlRenderText>(),
            text => text.Text.Contains("Footer", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(39)]
    [InlineData(64)]
    public void PagedRendererKeepsBorderedFlexContentWithItsFirstLine(int leadHeight) {
        const string html = "<style>@page{size:300px 100px;margin:0}html,body{margin:0}"
            + "body{max-width:200px;margin:0 auto}"
            + ".lead{height:LEADpx}nav{margin-top:32px;border:1px solid #ddd}"
            + "nav ul{display:flex;flex-wrap:wrap;margin:0;padding:8px;list-style:none}"
            + "nav li{flex:0 1 50%}nav a{display:flex;flex:1 1 100%}"
            + "nav span{display:flex;flex-direction:column}</style>"
            + "<a href='#pager'>Skip to navigation</a><div class='lead'></div><nav id='pager'><ul><li><a href='https://example.test/next'>"
            + "<span><span>Next:</span><span>One Header</span></span></a></li></ul></nav>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html.Replace("LEAD", leadHeight.ToString()),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Equal(2, rendered.Pages.Count);
        Assert.DoesNotContain(rendered.Pages[0].Visuals,
            visual => visual.Source == "nav#pager");
        HtmlRenderShape pager = Assert.Single(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(),
            visual => visual.Source == "nav#pager");
        Assert.Equal(0D, pager.Y, 3);
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderNamedDestination>(),
            destination => destination.Name == "pager" && Math.Abs(destination.Y) < 0.001D);
    }

    [Fact]
    public void PagedRendererKeepsCollapsedMarginBorderAtFragmentStart() {
        const string html = "<style>@page{size:300px 100px;margin:0}html,body{margin:0}"
            + "body{max-width:200px;margin:0 auto}"
            + ".lead{height:64px;margin-bottom:32px}nav{margin-top:32px;border:1px solid #ddd}"
            + "nav ul{display:flex;flex-wrap:wrap;margin:0;padding:8px;list-style:none}"
            + "nav li{flex:0 1 50%}nav a{display:flex;flex:1 1 100%}"
            + "nav span{display:flex;flex-direction:column}</style>"
            + "<div class='lead'></div><nav id='pager'><ul><li><a href='https://example.test/next'>"
            + "<span><span>Next:</span><span>One Header</span></span></a></li></ul></nav>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Equal(2, rendered.Pages.Count);
        Assert.DoesNotContain(rendered.Pages[0].Visuals,
            visual => visual.Source == "nav#pager");
        HtmlRenderShape pager = Assert.Single(rendered.Pages[1].Visuals.OfType<HtmlRenderShape>(),
            visual => visual.Source == "nav#pager");
        Assert.Equal(0D, pager.Y, 3);
    }

    [Fact]
    public void PagedRendererRetainsFlattenedTargetWhenFirstChildMovesToNextPage() {
        const string html = "<style>@page{size:300px 100px;margin:0}html,body{margin:0}"
            + "body{max-width:200px;margin:0 auto}.lead{height:64px}"
            + "section{display:contents}nav{margin-top:32px;border:1px solid #ddd}"
            + "nav ul{display:flex;margin:0;padding:8px;list-style:none}"
            + "nav a{display:flex}nav span{display:flex;flex-direction:column}</style>"
            + "<a href='#pager'>Skip</a><div class='lead'></div>"
            + "<section id='pager'><nav><ul><li><a href='https://example.test/next'>"
            + "<span><span>Next:</span><span>One Header</span></span></a></li></ul></nav></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderNamedDestination>(),
            destination => destination.Name == "pager" && Math.Abs(destination.Y) < 0.001D);
    }

    [Fact]
    public void PagedRendererRetainsTargetWithNegativeTopMargin() {
        const string html = "<style>@page{size:300px 100px;margin:0}html,body{margin:0}"
            + "p{break-before:page;margin-top:-8px}</style><a href='#target'>Jump</a>"
            + "<p id='target'>Target content</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderNamedDestination>(),
            destination => destination.Name == "target" && destination.Y >= -0.001D);
    }

    [Theory]
    [InlineData(false, 80, 0)]
    [InlineData(true, 80, 0)]
    [InlineData(true, 120, 24)]
    public void PagedRendererDiscardsOnlyUnpaintedParagraphMarginAtPageStart(bool insideFlex, int asideHeight, int targetY) {
        string opening = insideFlex ? "<div class='row'><div class='primary'>" : string.Empty;
        string closing = insideFlex ? "</div><div class='aside'>Aside</div></div>" : string.Empty;
        string html = "<style>@page{size:300px 100px;margin:0}html,body{margin:0;padding:0}"
            + "body{font:16px/20px Arial,sans-serif}.row{display:flex;gap:12px}"
            + ".primary{flex:1 1 auto;min-width:0}.aside{flex:0 0 72px;"
            + (asideHeight > 100 ? "min-height:" : "height:") + asideHeight + "px;background:#ddd}"
            + "p{margin:0 0 24px}#target{margin-bottom:0}</style>"
            + opening + "<p id='lead'>Lead one<br>Lead two<br>Lead three<br>Lead four</p>"
            + "<p id='target'>Target paragraph</p>" + closing;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, HonorCssPageRules = true, AutoFitWidePrintRoot = true });

        Assert.Equal(2, rendered.Pages.Count);
        foreach (string line in new[] { "Lead one", "Lead two", "Lead three", "Lead four" }) {
            Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == line);
            Assert.DoesNotContain(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == line);
        }
        HtmlRenderText target = Assert.Single(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(),
            text => text.Text.Contains("Target paragraph", StringComparison.Ordinal));
        Assert.InRange(target.Y, targetY, targetY + 2D);
        if (insideFlex && asideHeight > 100) {
            Assert.Contains(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderShape>(),
                shape => shape.Source == "div.aside" && shape.Y <= 0D && shape.Y + shape.Height >= 24D);
        }
    }

    [Fact]
    public void PagedRendererDoesNotDiscardFixedHeightFlowBesideParagraphMargin() {
        const string html = "<style>@page{size:300px 100px;margin:0}html,body{margin:0;padding:0}"
            + "body{font:16px/20px Arial}.row{display:flex}.tall{height:700px;width:150px}"
            + ".side{display:flow-root;width:150px}.side p{margin:0 0 200px}</style>"
            + "<div class='row'><div class='tall'>Tall</div><div class='side'><p>Short</p></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        // Preserve the baseline's authored-height flow; a sibling margin must
        // not compress this output to six pages by dropping blank box extent.
        Assert.Equal(8, rendered.Pages.Count);
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(),
            text => text.Text == "Tall");
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(),
            text => text.Text == "Short");
    }

    [Fact]
    public void PagedRendererPreservesRightPageBreakAfterParagraphMargin() {
        const string html = "<style>@page{size:300px 100px;margin:0}html,body{margin:0;padding:0}"
            + "body{font:16px/20px Arial}p{margin:0 0 24px}#lead{break-after:right}</style>"
            + "<main><p id='lead'>Lead one<br>Lead two<br>Lead three<br>Lead four</p>"
            + "<p id='target'>Target paragraph</p></main>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Equal(3, rendered.Pages.Count);
        Assert.DoesNotContain(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(),
            text => text.Text.Contains("Target paragraph", StringComparison.Ordinal));
        Assert.Contains(rendered.Pages[2].Visuals.OfType<HtmlRenderText>(),
            text => text.Text.Contains("Target paragraph", StringComparison.Ordinal));
    }

    [Fact]
    public void PagedRendererDoesNotResumeCompletedFlexParagraphAfterPageWidthChange() {
        const string html = "<style>@page{size:400px 100px;margin:0}@page:first{size:300px 100px}"
            + "html,body{margin:0;padding:0}body{font:16px/20px Arial}"
            + ".row{display:flex}.primary{flex:1;min-width:0}p{margin:0 0 24px}#target{margin-bottom:0}</style>"
            + "<main><div class='row'><div class='primary'>"
            + "<p>Lead one<br>Lead two<br>Lead three<br>Lead four</p>"
            + "<p id='target'>Target paragraph</p></div></div></main>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        foreach (string line in new[] { "Lead one", "Lead two", "Lead three", "Lead four", "Target paragraph" }) {
            Assert.Single(text, item => item.Text == line);
        }
    }

    [Fact]
    public void PagedFlexLinkKeepsPrintedUrlTogetherThroughHyphen() {
        const string html = "<style>@page{size:794px 140px;margin:0}html,body{margin:0;font:16px Arial}"
            + "nav{border:1px solid #ddd}ul{display:flex;margin:0;padding:8px;list-style:none}"
            + "li{display:flex;flex:0 1 100%}a{display:flex;flex:1 1 100%;flex-direction:row-reverse;justify-content:flex-end}"
            + "a::after{content:' (https://www.w3.org/WAI/tutorials/tables/one-header/)'}"
            + "a span{display:flex;flex:1 1 auto;width:100%;margin:0 8px;flex-direction:column}"
            + "a span span{margin:0}</style><nav><ul><li><a href='https://www.w3.org/WAI/tutorials/tables/one-header/'>"
            + "<span><span>Next:</span><span>One Header</span></span></a></li></ul></nav>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderText[] generated = EnumerateRenderVisuals(Assert.Single(rendered.Pages).Scene)
            .OfType<HtmlRenderText>()
            .Where(text => text.Source?.Contains("::after", StringComparison.Ordinal) == true)
            .ToArray();

        Assert.NotEmpty(generated);
        double firstLineY = generated.Min(text => text.Y);
        string firstLine = string.Concat(generated.Where(text => Math.Abs(text.Y - firstLineY) < 0.01D)
            .OrderBy(text => text.X).Select(text => text.Text));
        Assert.Contains("/one-", firstLine, StringComparison.Ordinal);
    }
}
