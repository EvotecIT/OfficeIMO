using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "", false)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "", false)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, "", false)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "z-index:2", false)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "opacity:.5", false)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "", true)]
    public void HtmlFloatPaint_NestedPositionedPanelKeepsDescendantGlyphsVisible(
        HtmlRenderIntentProfile profile, string context, bool clippedSemanticWrappers) {
        string html = "<style>html,body{margin:0;font:20px/26px Pinned;color:black}"
            + "#panel{width:300px;background:#ddd;position:relative;" + context + "}"
            + "ul{margin:0;padding:0}li{float:left;position:relative;list-style:none}"
            + "a{float:left}p{margin:0}table{background:lightblue}</style>"
            + (clippedSemanticWrappers ? "<main style='overflow:hidden'><section>" : "<div><div>")
            + "<div id='panel'><ul><li><a>JAN</a></li></ul><p>January:</p>"
            + "<table id='content'><tr><td>Visible cell</td></tr></table></div>"
            + (clippedSemanticWrappers ? "</section></main>" : "</div></div>");
        HtmlRenderOptions options = TableIntrinsicOptions();
        options.AllowSystemFontFallback = false;
        options.ViewportWidth = 320D;
        options.ViewportHeight = 300D;
        options.PageSize = new OfficePageSize(320D / HtmlRenderOptions.CssPixelsPerInch,
            300D / HtmlRenderOptions.CssPixelsPerInch);
        options.HonorCssPageRules = false;
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, options)).Document;

        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(page.Scene).ToArray();
        int panelBackground = Array.FindIndex(visuals,
            visual => visual is HtmlRenderShape && visual.Source == "div#panel");
        Assert.True(panelBackground >= 0);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing(), 1D, OfficeColor.White);
        foreach (string marker in new[] { "January:", "Visible" }) {
            HtmlRenderText text = Assert.Single(visuals.OfType<HtmlRenderText>(),
                item => item.Text.StartsWith(marker, StringComparison.Ordinal));
            Assert.True(panelBackground < Array.IndexOf(visuals, text),
                "An opaque ancestor background must paint before its descendant text.");
            Assert.True(FloatPaintContainsGlyphInk(raster, text),
                "Retained searchable text must also retain visible glyph ink.");
        }
        rendered.RequireNoLoss();
    }

    private static bool FloatPaintContainsGlyphInk(OfficeRasterImage raster, HtmlRenderText text) {
        int left = Math.Max(0, (int)Math.Floor(text.X));
        int right = Math.Min(raster.Width, (int)Math.Ceiling(text.X + (text.TextAdvanceWidth ?? text.Width)));
        int top = Math.Max(0, (int)Math.Floor(text.Y));
        int bottom = Math.Min(raster.Height, (int)Math.Ceiling(text.Y + text.Height));
        for (int y = top; y < bottom; y++) {
            for (int x = left; x < right; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.R < 150 && pixel.G < 150 && pixel.B < 150) return true;
            }
        }
        return false;
    }
}
