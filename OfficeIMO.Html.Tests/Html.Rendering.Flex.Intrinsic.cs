using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlFlexRow_InlineBlockPaddingContributesToNestedIntrinsicWidth() {
        HtmlRenderDocument rendered = RenderFlex("""
            <style>*{box-sizing:border-box}body{margin:0}</style>
            <div style="display:flex;width:300px">
              <div style="flex:1 1 auto">Released</div>
              <div style="flex:0 1 auto">
                <ul style="display:flex;gap:8px;list-style:none;margin:0;padding:0">
                  <li><span id="badge" style="display:inline-block;padding:2px 12px;background:#888888">ID: 40550</span></li>
                  <li style="width:24px;flex:none">X</li>
                </ul>
              </div>
            </div>
            """, 300D);

        HtmlRenderText[] badgeText = rendered.Pages[0].Visuals.OfType<HtmlRenderText>()
            .Where(text => text.Source == "span")
            .ToArray();
        Assert.NotEmpty(badgeText);
        Assert.Single(badgeText.Select(text => Math.Round(text.Y, 3)).Distinct());
        HtmlRenderShape badge = FindFlexShape(rendered, "span#badge");
        Assert.True(badge.Width > 60D, "the inline-block text and both padding edges must fit");
    }

    [Fact]
    public void HtmlFlexRow_PercentageInlineBlockDoesNotDefineAutoParentWidth() {
        HtmlRenderDocument rendered = RenderFlex("""
            <div style="display:flex;width:300px">
              <div id="auto" style="flex:0 0 auto;background:#ff0000">
                <span style="display:inline-block;width:100%">ID</span>
              </div>
              <div id="next" style="flex:0 0 50px;background:#00ff00">Next</div>
            </div>
            """, 300D);

        HtmlRenderShape auto = FindFlexShape(rendered, "div#auto");
        HtmlRenderShape next = FindFlexShape(rendered, "div#next");
        Assert.True(auto.Width < 100D, "the cyclic percentage must not consume the flex row");
        Assert.Equal(auto.X + auto.Width, next.X, 1);
    }

    [Fact]
    public void HtmlFlexRow_NestedInlineBlocksRespectLayoutDepthLimit() {
        string html = "<div style='display:flex'>" + string.Concat(Enumerable.Repeat("<span style='display:inline-block'>", 12))
            + "text" + string.Concat(Enumerable.Repeat("</span>", 12)) + "</div>";

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { MaxLayoutDepth = 8 }));

        Assert.Equal(HtmlRenderDiagnosticCodes.DepthLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutDepth), exception.LimitSource);
    }
}
