using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("none", 1D)]
    [InlineData("none", 0.5D)]
    [InlineData("perspective(20px)", 1D)]
    [InlineData("perspective(20px)", 0.5D)]
    public void HtmlAbsoluteWrapper_PreservesNormalFlowChildPaintTransform(string wrapperTransform, double opacity) {
        string html = "<div style='position:relative;width:30px;height:160px;margin:0'>"
            + "<div style='position:absolute;left:0;top:0;width:30px;height:80px;margin:0;transform:"
            + wrapperTransform + "'>"
            + "<div style='width:30px;height:20px;margin:0;background:red;transform:translateY(60px);opacity:"
            + opacity.ToString(System.Globalization.CultureInfo.InvariantCulture) + "'></div></div></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(40D / HtmlRenderOptions.CssPixelsPerInch, 80D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(2, rendered.Pages.Count);
        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeRasterImage second = OfficeDrawingRasterRenderer.Render(rendered.Pages[1].CreateDrawing());
        OfficeColor painted = first.GetPixel(10, 70);
        Assert.Equal(255, painted.R);
        Assert.Equal(0, painted.G);
        Assert.Equal(0, painted.B);
        Assert.InRange((int)painted.A, (int)(opacity * 255D), (int)Math.Ceiling(opacity * 255D));
        Assert.Equal(OfficeColor.Transparent, first.GetPixel(10, 10));
        Assert.Equal(OfficeColor.Transparent, second.GetPixel(10, 10));
    }

    [Theory]
    [InlineData(false, 80D)]
    [InlineData(true, 80D)]
    [InlineData(false, 70D)]
    [InlineData(true, 70D)]
    public void HtmlAbsoluteRows_PreservePaintAcrossPagedContainingBlock(bool useTranslation, double pageHeight) {
        // A virtualized report retains a tall containing block and individually
        // positioned rows. Inspect real paint: logical text can survive off-page.
        string rows = string.Concat(Enumerable.Range(0, 8).Select(index => {
            string position = useTranslation
                ? "top:0;transform:translateY(" + (index * 20) + "px)"
                : "top:" + (index * 20) + "px";
            string color = index % 2 == 0 ? "#ff0000" : "#0000ff";
            return "<div style='position:absolute;left:0;width:30px;height:20px;margin:0;background:"
                + color + ";" + position + "'></div>";
        }));
        string html = "<div style='position:relative;width:30px;height:160px;margin:0'>" + rows + "</div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(40D / HtmlRenderOptions.CssPixelsPerInch, pageHeight / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        int rowsPerPage = (int)(pageHeight / 20D);
        Assert.Equal((8 + rowsPerPage - 1) / rowsPerPage, rendered.Pages.Count);
        for (int pageIndex = 0; pageIndex < rendered.Pages.Count; pageIndex++) {
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[pageIndex].CreateDrawing());
            int firstRow = pageIndex * rowsPerPage;
            for (int localRow = 0; localRow < Math.Min(rowsPerPage, 8 - firstRow); localRow++) {
                OfficeColor expected = (firstRow + localRow) % 2 == 0 ? OfficeColor.Red : OfficeColor.Blue;
                Assert.Equal(expected, raster.GetPixel(10, localRow * 20 + 10));
            }
        }
    }
}
