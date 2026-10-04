using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderIntentProfile.ScreenFullPage)]
    [InlineData(HtmlRenderIntentProfile.ScreenViewport)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    public void HtmlRender_Screen_PaintsDistinctBodyBackgroundInsideHtmlCanvas(HtmlRenderIntentProfile profile) {
        const string html = """
            <style>
              html { background: #222; }
              body { margin: 0; background: white; color: #17171b; }
              header, footer { height: 40px; background: black; color: white; }
              main { height: 160px; }
            </style>
            <header>Header marker</header><main>Readable article marker</main><footer>Footer marker</footer>
            """;
        var options = new HtmlRenderOptions {
            ViewportWidth = 384D,
            ViewportHeight = 288D,
            PageSize = new OfficePageSize(4D, 1.5D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false
        };
        HtmlRenderResult result = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, options));

        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(result.Document.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.White, first.GetPixel(370, 100));
        OfficeRasterImage last = OfficeDrawingRasterRenderer.Render(result.Document.Pages[^1].CreateDrawing(), 1D, OfficeColor.White);
        int offset = profile == HtmlRenderIntentProfile.ScreenSnapshotPaged ? 144 : 0;
        Assert.Equal(OfficeColor.Black, last.GetPixel(370, 220 - offset));
        Assert.Equal(OfficeColor.FromRgb(0x22, 0x22, 0x22), last.GetPixel(370, 270 - offset));
        Assert.Contains("Readable article marker", result.Document.Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderUserAgentStyleMode.Document)]
    [InlineData(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderUserAgentStyleMode.Browser)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderUserAgentStyleMode.Document)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, HtmlRenderUserAgentStyleMode.Document)]
    public void HtmlRender_ContentsBodyKeepsChildrenWithoutPaintingABodyBox(
        HtmlRenderIntentProfile profile, HtmlRenderUserAgentStyleMode userAgentStyles) {
        const string html = """
            <style>
              html { background: #222; }
              body { margin: 0; background: white; color: white; }
              section { height: 80px; }
            </style>
            <body style="display:contents"><section>Contents child marker</section></body>
            """;
        var options = new HtmlRenderOptions {
            ViewportWidth = 384D,
            ViewportHeight = 288D,
            PageSize = new OfficePageSize(4D, 1.5D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false,
            UserAgentStyles = userAgentStyles
        };
        HtmlRenderResult result = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, options));

        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(result.Document.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.FromRgb(0x22, 0x22, 0x22), image.GetPixel(370, 60));
        Assert.Contains("Contents child marker", result.Document.Text, StringComparison.Ordinal);
    }
}
