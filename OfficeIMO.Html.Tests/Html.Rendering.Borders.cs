using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlBorders_UniformRadiusFlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<div id='rounded' style='width:40px;height:20px;margin:0;border:2px solid #0000ff;border-radius:6px;background:#ff0000;font-size:6px;line-height:8px'>RoundedPdf</div>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 50D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        List<HtmlRenderShape> shapes = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == "div#rounded")
            .ToList();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(50D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = html.SaveAsPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Load(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(2, shapes.Count);
        Assert.All(shapes, shape => {
            Assert.Equal(OfficeShapeKind.RoundedRectangle, shape.Shape.Kind);
            Assert.Equal(6D, shape.Shape.CornerRadius, 3);
        });
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(0, 0));
        Assert.Equal((byte)255, raster.GetPixel(20, 15).R);
        Assert.Contains("rx=\"6\" ry=\"6\"", svg, StringComparison.Ordinal);
        Assert.Contains("RoundedPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlBorders_AsymmetricRadiusUsesCatalogedSquareFallback() {
        const string html = "<div id='asymmetric' style='width:30px;height:20px;margin:0;border-radius:8px 2px;background:#ff0000'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        HtmlRenderShape shape = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), item => item.Source == "div#asymmetric");
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported);

        Assert.Equal(OfficeShapeKind.Rectangle, shape.Shape.Kind);
        Assert.Equal("div#asymmetric", diagnostic.Source);
        Assert.Contains("asymmetric-or-elliptical", diagnostic.Detail, StringComparison.Ordinal);
        Assert.Contains(HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-radius:6px)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-top-left-radius:6px)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border-radius:6px 2px)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border-top-left-radius:6px 2px)"));
    }
}
