using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void AbsolutePercentageWidth_UsesResolvedGridAreaWidth() {
        const string html = "<div style='display:grid;grid-template-columns:100px 200px;grid-template-rows:100px;width:300px'>"
            + "<div id='overlay' style='position:absolute;grid-column:2;grid-row:1;width:100%;height:20px;background:red'>Overlay</div>"
            + "</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderShape overlay = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderShape>()
            .First(shape => shape.Source == "div#overlay");
        Assert.Equal(200D, overlay.Width, 1);
    }
}
