using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, HtmlRenderMode.Continuous)]
    [InlineData(false, HtmlRenderMode.Paged)]
    [InlineData(true, HtmlRenderMode.Paged)]
    public void HtmlMargins_OverlappingBlockBackgroundKeepsEarlierInFlowGlyphsVisible(
        bool columns, HtmlRenderMode mode) {
        string wrapper = columns
            ? "width:220px;height:100px;column-count:2;column-gap:20px;column-fill:auto"
            : "width:200px;padding:10px";
        string html = "<!doctype html><style>html,body{margin:0;font:14px/20px Pinned}"
            + "a{color:inherit;text-decoration:none}</style><section style='" + wrapper + "'>"
            + "<div style='height:20px'></div>"
            + "<div style='height:40px;color:red'><a href='#tail'>FIRST</a></div>"
            + "<div id='cover' style='height:25px;margin-top:-45px;background:#00aa00;color:blue'>OVER</div>"
            + "<div id='tail' style='height:20px'>TAIL</div></section>";
        HtmlRenderOptions options = TableIntrinsicOptions();
        options.Mode = mode;
        options.AllowSystemFontFallback = false;
        options.ViewportWidth = 320D;
        options.ViewportHeight = 100D;
        options.PageSize = new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(320D / HtmlRenderOptions.CssPixelsPerInch),
            global::OfficeIMO.OfficeLength.Inches(100D / HtmlRenderOptions.CssPixelsPerInch));
        options.HonorCssPageRules = false;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(page.Scene).ToArray();
        HtmlRenderText first = Assert.Single(visuals.OfType<HtmlRenderText>(), text => text.Text == "FIRST");
        HtmlRenderText over = Assert.Single(visuals.OfType<HtmlRenderText>(), text => text.Text == "OVER");
        HtmlRenderText tail = Assert.Single(visuals.OfType<HtmlRenderText>(), text => text.Text == "TAIL");
        OfficeColor green = OfficeColor.FromRgb(0, 170, 0);
        HtmlRenderShape background = Assert.Single(visuals.OfType<HtmlRenderShape>(), shape =>
            shape.Source == "div#cover" && shape.Shape.FillColor == green);
        Assert.Equal(columns ? 15D : 25D, background.Y, 4);
        Assert.Equal(first.Y + 20D, tail.Y, 4);
        Assert.Equal("#tail", first.LinkUri);
        Assert.Equal(OfficeColor.Red, first.Color);
        Assert.Equal(OfficeColor.Blue, over.Color);
        Assert.True(Array.IndexOf(visuals, background) < Array.IndexOf(visuals, first),
            "An ordinary overlapping block background must paint below earlier in-flow text.");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing(), 1D, OfficeColor.White);
        Assert.True(NegativeFlowContainsRedInk(raster, first, background),
            "Retained earlier text must paint visible red glyph ink within the later block's green background.");
        Assert.Equal(green, raster.GetPixel((int)(background.X + background.Width - 2D),
            (int)(background.Y + 2D)));
        rendered.RequireNoLoss();
    }

    private static bool NegativeFlowContainsRedInk(OfficeRasterImage raster, HtmlRenderText text,
        HtmlRenderShape background) {
        int left = Math.Max(0, (int)Math.Ceiling(Math.Max(text.X, background.X)));
        int right = Math.Min(raster.Width, (int)Math.Floor(Math.Min(
            text.X + (text.TextAdvanceWidth ?? text.Width), background.X + background.Width)));
        int top = Math.Max(0, (int)Math.Ceiling(Math.Max(text.Y, background.Y)));
        int bottom = Math.Min(raster.Height, (int)Math.Floor(Math.Min(
            text.Y + text.Height, background.Y + background.Height)));
        for (int y = top; y < bottom; y++) {
            for (int x = left; x < right; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.R > 180 && pixel.G < 100 && pixel.B < 100) return true;
            }
        }
        return false;
    }
}
