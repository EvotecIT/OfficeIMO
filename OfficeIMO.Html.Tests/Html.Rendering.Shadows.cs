using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlShadows_OuterBlurFlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<div id='shadow' style='width:28px;height:16px;margin:4px 0 0 8px;border-radius:4px;background:#ffffff;box-shadow:4px 3px 4px rgba(255,0,0,.5);font-size:6px;line-height:8px'>ShadowPdf</div>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 50D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderShape carrier = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#shadow:box-shadow");
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
        string rawPdf = Encoding.ASCII.GetString(pdf);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Load(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.NotNull(carrier.Shape.Shadow);
        Assert.Equal(4D, carrier.Shape.Shadow!.BlurRadius, 3);
        Assert.Equal(128D / 255D, carrier.Shape.Shadow.Opacity, 3);
        Assert.True(raster.GetPixel(38, 12).R > raster.GetPixel(38, 12).B);
        Assert.True(raster.GetPixel(38, 12).A > 0);
        Assert.Contains("fill=\"#FF0000\"", svg, StringComparison.Ordinal);
        Assert.Contains("/Type /ExtGState", rawPdf, StringComparison.Ordinal);
        Assert.Contains("ShadowPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlShadows_UnsupportedFormsUseCatalogedDiagnostics() {
        const string html = "<div id='inset-shadow' style='width:20px;height:20px;box-shadow:inset 0 1px 2px black'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 30D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported);

        Assert.Equal("div#inset-shadow", diagnostic.Source);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#inset-shadow:box-shadow");
        Assert.Contains(HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(box-shadow:0 1px 2px rgba(0,0,0,.2))"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(box-shadow:0 1px 2px 0 black)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(box-shadow:inset 0 1px 2px black)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(box-shadow:0 1px 2px 3px black)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(box-shadow:0 1px black, 0 2px blue)"));
    }
}
