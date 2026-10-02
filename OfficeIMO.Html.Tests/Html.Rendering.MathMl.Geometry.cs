using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlMathMl_InlineScopedTextSharesItsNeighborsPaintedBaseline() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<body style='margin:0;font:20px ScopedMath;line-height:24px'><p style='margin:0'>"
            + "x<math id='plain' style='font-family:inherit'><mtext>x</mtext></math>x</p></body>", ScopedMathRenderOptions());
        HtmlRenderDrawing math = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        OfficeDrawingText glyph = Assert.Single(math.Drawing.Elements.OfType<OfficeDrawingText>());
        double baseline = math.Y + glyph.Y + glyph.Font.Size + glyph.BaselineOffset;
        Assert.All(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), neighbor =>
            Assert.Equal(baseline, neighbor.Y + neighbor.Font.Size + neighbor.BaselineOffset, 6));
    }

    [Fact]
    public void HtmlMathMl_NestedFractionsDoNotShrinkTheWholeExpression() {
        const string html = "<body style='margin:0;font:20px ScopedMath;line-height:24px'>"
            + "<math id='simple' style='font-family:inherit'><mfrac><mtext>x</mtext><mn>2</mn></mfrac></math>"
            + "<math id='nested' style='font-family:inherit'><mfrac><mfrac><mtext>x</mtext><mn>2</mn></mfrac><mn>2</mn></mfrac></math></body>";
        var options = ScopedMathRenderOptions();
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderDrawing simple = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>(), item => item.Source == "math#simple");
        HtmlRenderDrawing nested = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>(), item => item.Source == "math#nested");

        Assert.Equal(1D, simple.Height / simple.Drawing.Height, 6);
        Assert.Equal(1D, nested.Height / nested.Drawing.Height, 6);
        Assert.Equal(14.2D, Assert.Single(simple.Drawing.Elements.OfType<OfficeDrawingText>(), item => item.Text == "x").Font.Size, 6);
        Assert.Equal(10.082D, Assert.Single(nested.Drawing.Elements.OfType<OfficeDrawingText>(), item => item.Text == "x").Font.Size, 6);
        Assert.True(nested.Height > simple.Height);
        Assert.Equal("(x)/(2)\n((x)/(2))/(2)", rendered.Text);
    }

    [Theory]
    [InlineData("center", 0.5D)]
    [InlineData("right", 1D)]
    public void HtmlMathMl_BlockBackgroundAndBorderFollowAlignedContent(string alignment, double fraction) {
        string html = "<body style='margin:0;font:20px ScopedMath'><math id='aligned' display='block' "
            + "style='font-family:inherit;text-align:" + alignment + ";background:#ff0000;border:2px solid blue;padding:3px'>"
            + "<mtext>x</mtext></math></body>";
        var options = ScopedMathRenderOptions();
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderDrawing drawing = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        HtmlRenderShape fill = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), item =>
            item.Source == "math#aligned" && item.Shape.FillColor == OfficeColor.Red);
        double expectedBoxX = (options.ViewportWidth - fill.Width) * fraction;
        Assert.Equal(expectedBoxX, fill.X, 6);
        Assert.Equal(fill.X + 5D, drawing.X, 6);
        Assert.Equal(fill.Y + 5D, drawing.Y, 6);
        Assert.InRange(drawing.X + drawing.Width, fill.X, fill.X + fill.Width);
    }

    [Fact]
    public void HtmlMathMlPdf_ScopedCompactGlyphRetainsItsFullInkAndLogicalText() {
        const string html = "<html lang='en'><body style='margin:0;font:20px ScopedMath'>"
            + "<math aria-label='x over two' style='font-family:inherit'><mfrac><mtext>x</mtext><mn>2</mn></mfrac></math></body></html>";
        var options = new HtmlToPdfOptions {
            PageSize = new OfficePageSize(2D, 1D), HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D), BackgroundColor = OfficeColor.Transparent,
            AllowSystemFontFallback = false
        };
        options.Fonts.Add("ScopedMath", ManagedTextShapingTestAssets.CreateFont('x', '2'));
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        Assert.Contains("(x)/(2)", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(pdf));
        OfficeDrawing reopened = PdfCore.PdfPageImageRenderer.RenderPage(pdf);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(reopened, scale: 4D);
        var rows = new SortedSet<int>();
        for (int y = 0; y < raster.Height; y++) {
            for (int x = 0; x < raster.Width; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.A > 127 && pixel.R < 127 && pixel.G < 127 && pixel.B < 127) rows.Add(y);
            }
        }
        // Each 14.2px glyph paints a 9.94px-high rectangle (29.82px at this PDF raster scale).
        // A baseline offset incorrectly applied to the PDF clip removes most of these rows.
        Assert.True(rows.Count >= 58, $"Expected both compact glyphs to retain their complete ink; observed {rows.Count} painted rows.");
    }

    [Fact]
    public void HtmlMathMlPdf_FractionGlyphsRetainTheirSharedHorizontalCenter() {
        const string html = "<body style='margin:0;font:20px ScopedMath'><math display='block' style='font-family:inherit'>"
            + "<mfrac><mtext>x</mtext><mn>22</mn></mfrac></math></body>";
        var options = new HtmlToPdfOptions {
            PageSize = new OfficePageSize(2D, 2D), HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D), BackgroundColor = OfficeColor.Transparent,
            AllowSystemFontFallback = false
        };
        options.Fonts.Add("ScopedMath", ManagedTextShapingTestAssets.CreateFont('x', '2'));
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfCore.PdfPageImageRenderer.RenderPage(pdf), scale: 4D);
        var groups = new List<(int Left, int Right, int Rows)>();
        bool active = false;
        for (int y = 0; y < raster.Height; y++) {
            int left = raster.Width, right = -1;
            for (int x = 0; x < raster.Width; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.A > 127 && pixel.R < 127) { left = Math.Min(left, x); right = Math.Max(right, x); }
            }
            if (right < 0) { active = false; continue; }
            if (!active) { groups.Add((left, right, 1)); active = true; }
            else { var prior = groups[groups.Count - 1]; groups[groups.Count - 1] = (Math.Min(prior.Left, left), Math.Max(prior.Right, right), prior.Rows + 1); }
        }
        var glyphs = groups.Where(group => group.Rows > 10).ToArray();
        Assert.Equal(2, glyphs.Length);
        // Fixture glyph ink is inset equally within each natural advance. The numerator
        // and denominator therefore share an ink center, independently of their widths.
        Assert.InRange(Math.Abs((glyphs[0].Left + glyphs[0].Right) - (glyphs[1].Left + glyphs[1].Right)), 0, 2);
    }

    [Fact]
    public void HtmlMathMlPdf_SyntheticStyleRetainsTheOwnedPositionedInk() {
        const string html = "<body style='margin:0;font:italic bold 144px ScopedMath'><math style='font-family:inherit;font-weight:inherit;font-style:inherit'><mtext>x</mtext></math></body>";
        var renderOptions = ScopedMathRenderOptions();
        renderOptions.ViewportWidth = 384D; renderOptions.ViewportHeight = 384D;
        HtmlRenderDrawing math = Assert.Single(HtmlRenderTestDriver.Render(html, renderOptions).Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        var pdfOptions = new HtmlToPdfOptions {
            PageSize = new OfficePageSize(4D, 4D), HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D), BackgroundColor = OfficeColor.Transparent,
            AllowSystemFontFallback = false
        };
        pdfOptions.Fonts.Add("ScopedMath", ManagedTextShapingTestAssets.CreateFont('x', '2'));
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(pdfOptions);
        OfficeRasterImage pdfRaster = OfficeDrawingRasterRenderer.Render(PdfCore.PdfPageImageRenderer.RenderPage(pdf), scale: 4D);
        OfficeRasterImage nativeRaster = OfficeDrawingRasterRenderer.Render(math.Drawing, scale: 3D);
        int expected = CountBlackInk(nativeRaster), actual = CountBlackInk(pdfRaster);
        // Four PDF pixels per point equal three pixels per CSS drawing unit. The
        // independent reread must retain synthetic bold ink from the regular-only face.
        Assert.InRange(actual, (int)(expected * .98D), (int)(expected * 1.02D));
    }

    private static int CountBlackInk(OfficeRasterImage raster) {
        int count = 0;
        for (int y = 0; y < raster.Height; y++)
            for (int x = 0; x < raster.Width; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.A > 127 && pixel.R < 127) count++;
            }
        return count;
    }

    private static HtmlRenderOptions ScopedMathRenderOptions() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 360D, ViewportHeight = 160D, Margins = HtmlRenderMargins.All(0D),
            AllowSystemFontFallback = false
        };
        options.Fonts.Add("ScopedMath", ManagedTextShapingTestAssets.CreateFont('x', '2'));
        return options;
    }
}
