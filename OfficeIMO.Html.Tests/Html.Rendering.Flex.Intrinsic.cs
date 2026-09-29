using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlInlineFlex_MaxContentWidthFitsTextMeasuredByTokens() {
        const string labelText = "Page Last Updated:";
        const string html = """
            <style>
              body { margin: 0; font-size: 16px; }
              li { display: inline-flex; font-family: 'OfficeIMO Shaping Test'; font-size: .9rem; }
              #label { padding-right: 8px; }
            </style>
            <li><div id="label">Page Last Updated:</div><div>Jan 28, 2026</div></li>
            """;
        var options = new HtmlRenderOptions {
            ViewportWidth = 768D,
            Margins = HtmlRenderMargins.All(0D),
            TextShapingProvider = new IntrinsicWidthShaper()
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs((labelText + "Jan 28, 2026")
                .Distinct().Select(character => (int)character).ToArray()));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderText[] label = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>()
            .Where(text => text.Text.Contains("Page", StringComparison.Ordinal)
                || text.Text.Contains("Updated", StringComparison.Ordinal))
            .ToArray();
        Assert.NotEmpty(label);
        Assert.Single(label.Select(text => Math.Round(text.Y, 3)).Distinct());
    }

    private sealed class IntrinsicWidthShaper : IOfficeTextShapingProvider {
        public OfficeTextShapingResult ShapeText(OfficeTextShapingRequest request) {
            var glyphs = new List<OfficeShapedGlyph>();
            int index = 0;
            bool first = true;
            foreach (string element in OfficeTextElements.Enumerate(request.Text)) {
                int advance = first && request.Text.Trim().Contains(' ') ? 490 : 500;
                glyphs.Add(new OfficeShapedGlyph(1, element, index, advance));
                first = false;
                index += element.Length;
            }
            return new OfficeTextShapingResult(glyphs);
        }
    }

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
