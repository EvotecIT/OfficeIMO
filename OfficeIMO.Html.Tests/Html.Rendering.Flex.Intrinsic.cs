using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
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

    [Theory]
    [InlineData("block", "100%", "0", "0", 242.909D)]
    [InlineData("inline-block", "100%", "0", "0", 242.909D)]
    [InlineData("block", "300px", "0", "0", 300D)]
    [InlineData("block", "100%", "280px", "0", 280D)]
    [InlineData("block", "100%", "0", "200px", 242.909D)]
    public void HtmlFlexRow_FixedWidthImageWithPercentageMaximumCanShrink(
        string wrapperDisplay, string maximumWidth, string minimumWidth, string minimumHeight, double expectedSidebarWidth) {
        string image = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(400, 200));
        string html = "<style>body{margin:0}</style><div style='display:flex;width:700px'>"
            + "<div id='prose' style='width:100%;margin-right:32px;background:#eeeeee'>Prose</div>"
            + "<div id='sidebar' style='background:#ddeeff'><div style='width:100%;max-width:100%'>"
            + "<figure style='margin:0;display:" + wrapperDisplay + "'><img src='data:image/png;base64," + image
            + "' style='display:block;width:400px;max-width:" + maximumWidth + ";min-width:" + minimumWidth
            + ";height:auto;min-height:" + minimumHeight + "'></figure></div></div></div>";

        HtmlRenderDocument rendered = RenderFlex(html, 700D);
        HtmlRenderShape prose = FindFlexShape(rendered, "div#prose");
        HtmlRenderShape sidebar = FindFlexShape(rendered, "div#sidebar");
        HtmlRenderImage renderedImage = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());

        Assert.Equal(expectedSidebarWidth, sidebar.Width, 1);
        Assert.Equal(668D - expectedSidebarWidth, prose.Width, 1);
        Assert.Equal(prose.X + prose.Width + 32D, sidebar.X, 1);
        Assert.Equal(sidebar.Width, renderedImage.Width, 1);
        Assert.Equal(minimumHeight == "200px" ? 200D : renderedImage.Width / 2D, renderedImage.Height, 1);
        Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Prose");
    }

    [Theory]
    [InlineData("width:400px;max-width:100%;min-width:50%", 242.909D)]
    [InlineData("width:400px;max-width:calc(50% + 300px)", 400D)]
    [InlineData("width:calc(50% + 300px);max-width:none", 421.455D)]
    public void HtmlFlexRow_CyclicImageConstraintsUseIntrinsicReferenceForMinimum(string constraints, double expectedImageWidth) {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(400, 200));
        string html = "<style>body{margin:0}</style><div style='display:flex;width:700px'>"
            + "<div id='prose' style='width:100%;margin-right:32px;background:#eeeeee'>Prose</div>"
            + "<div id='sidebar' style='background:#ddeeff'><figure style='margin:0'><img style='display:block;height:auto;"
            + constraints + "' src='data:image/png;base64," + data + "'></figure></div></div>";

        HtmlRenderDocument rendered = RenderFlex(html, 700D);
        HtmlRenderShape sidebar = FindFlexShape(rendered, "div#sidebar");
        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(242.909D, sidebar.Width, 1);
        Assert.Equal(expectedImageWidth, image.Width, 1);
        Assert.Equal(sidebar.X, image.X, 1);
        Assert.Equal(expectedImageWidth / 2D, image.Height, 1);
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
