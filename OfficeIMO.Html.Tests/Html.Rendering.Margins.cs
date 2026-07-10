using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlMargins_AdjacentBlockMarginsCollapseWithNegativeValuesAcrossBackends() {
        const string html = "<div id='container' style='width:50px;margin:0'>"
            + "<div id='first' style='width:20px;height:10px;margin:0 0 20px;background:#ff0000'></div>"
            + "<div id='second' style='width:20px;height:10px;margin:-5px 0 8px;background:#0000ff'></div>"
            + "<div id='third' style='width:40px;height:10px;margin:12px 0 0;background:#00ff00;font-size:6px;line-height:8px'>GapPdf</div>"
            + "</div>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 55D,
            ViewportHeight = 58D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderShape first = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#first");
        HtmlRenderShape second = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#second");
        HtmlRenderShape third = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#third");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(55D / HtmlRenderOptions.CssPixelsPerInch, 58D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = html.SaveAsPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Load(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(0D, first.Y, 3);
        Assert.Equal(25D, second.Y, 3);
        Assert.Equal(47D, third.Y, 3);
        Assert.Equal(57D, third.Y + third.Height, 3);
        Assert.True(raster.GetPixel(5, 25).B > raster.GetPixel(5, 25).R);
        Assert.Contains("y=\"25\"", svg, StringComparison.Ordinal);
        Assert.Contains("GapPdf", pdfText, StringComparison.Ordinal);
        Assert.Single(PdfCore.PdfReadDocument.Load(pdf).Pages);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }
}
