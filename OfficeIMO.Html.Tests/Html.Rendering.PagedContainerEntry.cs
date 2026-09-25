using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void PagedRendererKeepsBorderedFlexContentWithItsFirstLine() {
        const string html = "<style>@page{size:300px 100px;margin:0}html,body{margin:0}"
            + "body{max-width:200px;margin:0 auto}"
            + ".lead{height:64px}nav{margin-top:32px;border:1px solid #ddd}"
            + "nav ul{display:flex;flex-wrap:wrap;margin:0;padding:8px;list-style:none}"
            + "nav li{flex:0 1 50%}nav a{display:flex;flex:1 1 100%}"
            + "nav span{display:flex;flex-direction:column}</style>"
            + "<a href='#pager'>Skip to navigation</a><div class='lead'></div><nav id='pager'><ul><li><a href='https://example.test/next'>"
            + "<span><span>Next:</span><span>One Header</span></span></a></li></ul></nav>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
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
}
