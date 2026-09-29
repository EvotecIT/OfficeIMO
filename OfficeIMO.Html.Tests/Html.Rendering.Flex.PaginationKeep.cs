using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlFlexRow_KeepsFittingAutoHeightLineOnNextPage() {
        const string html = """
            <style>body,p{margin:0}</style>
            <div style="height:50px">Before</div>
            <div id="row" style="display:flex;width:160px;align-items:flex-start">
              <div style="width:80px;line-height:20px">First<br>Second<br>Third<br>Fourth</div>
              <div style="width:80px;height:20px">Side</div>
            </div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 100D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text =>
            text.Text is "First" or "Second" or "Third" or "Fourth" or "Side");
        Assert.All(new[] { "First", "Second", "Third", "Fourth", "Side" }, marker =>
            Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == marker));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
    }
}
