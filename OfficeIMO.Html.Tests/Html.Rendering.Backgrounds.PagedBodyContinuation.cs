using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_PagedBodyBackgroundCoversSlackBeforeContinuation() {
        const string html = """
            <style>
              @page { size: 400px 400px; margin: 0; }
              html { background: #222; }
              body { margin: 0; display: flex; flex-direction: column; min-height: 100vh; background: #eee; }
              section { height: 300px; flex: none; }
              section:first-of-type { background: #ddd; }
            </style>
            <section>First block</section><section>Second block</section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            HonorCssPageRules = true,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Equal(2, rendered.Pages.Count);
        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        OfficeRasterImage last = OfficeDrawingRasterRenderer.Render(rendered.Pages[1].CreateDrawing(), 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.FromRgb(0xee, 0xee, 0xee), first.GetPixel(200, 350));
        Assert.Equal(OfficeColor.FromRgb(0x22, 0x22, 0x22), last.GetPixel(200, 350));
    }
}
