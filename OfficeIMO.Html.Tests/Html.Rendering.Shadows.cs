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
        var options = new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape carrier = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#shadow:box-shadow");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(50D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string rawPdf = Encoding.ASCII.GetString(pdf);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.NotNull(carrier.Shape.Shadow);
        Assert.Equal(4D, carrier.Shape.Shadow!.BlurRadius, 3);
        Assert.Equal(128D / 255D, carrier.Shape.Shadow.Opacity, 3);
        Assert.True(raster.GetPixel(38, 12).R > raster.GetPixel(38, 12).B);
        Assert.True(raster.GetPixel(38, 12).A > 0);
        Assert.Contains("fill=\"#FF0000\"", svg, StringComparison.Ordinal);
        Assert.Contains("/Type /ExtGState", rawPdf, StringComparison.Ordinal);
        Assert.Contains("ShadowPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlShadows_MultipleOuterLayersAndSignedSpreadUseCssPaintOrder() {
        const string html = "<body style='margin:0'><div id='shadow' style='width:20px;height:10px;margin:10px;background:white;box-shadow:0 0 0 3px red,4px 0 0 -1px blue'></div></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 45D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape[] shadows = rendered.Pages[0].Visuals
            .OfType<HtmlRenderShape>()
            .Where(shape => shape.Source?.StartsWith("div#shadow:box-shadow[", StringComparison.Ordinal) == true)
            .ToArray();
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        Assert.Equal(2, shadows.Length);
        HtmlRenderShape red = Assert.Single(shadows, shape => shape.Source == "div#shadow:box-shadow[0]");
        HtmlRenderShape blue = Assert.Single(shadows, shape => shape.Source == "div#shadow:box-shadow[1]");
        Assert.Equal(7D, red.X, 3);
        Assert.Equal(26D, red.Width, 3);
        Assert.Equal(11D, blue.X, 3);
        Assert.Equal(18D, blue.Width, 3);
        Assert.True(blue.PaintOrder < red.PaintOrder);
        Assert.Equal(4D, blue.Shape.Shadow!.OffsetX, 3);
        Assert.Contains("fill=\"#FF0000\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#0000FF\"", svg, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported);
    }

    [Fact]
    public void HtmlShadows_InsetSpreadAndBlurAreClippedAcrossPngAndSvg() {
        const string html = "<body style='margin:0'><div id='inset-shadow' style='width:30px;height:20px;margin:5px;border-radius:4px;background:white;box-shadow:inset 3px 0 2px 2px rgba(255,0,0,.8)'></div></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderPathClipGroup group = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderPathClipGroup>(),
            item => item.Source == "div#inset-shadow:box-shadow[0]:inset");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(40D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);

        Assert.Equal(5, group.Visuals.Count);
        Assert.True(raster.GetPixel(7, 15).R > raster.GetPixel(7, 15).G);
        Assert.True(raster.GetPixel(20, 15).G > 220);
        Assert.Contains("fill-rule=\"evenodd\"", svg, StringComparison.Ordinal);
        Assert.Contains("clip-path=", svg, StringComparison.Ordinal);
        Assert.Equal(1, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.Contains("/Type /ExtGState", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported);
    }

    [Fact]
    public void HtmlShadows_LayerLimitIsBoundedAndCataloged() {
        const string html = "<div id='limited' style='width:20px;height:20px;box-shadow:0 1px red,0 2px green,0 3px blue'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D),
            MaxBoxShadowLayers = 2
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.BoxShadowLayerLimit);

        Assert.Equal(2, rendered.Pages[0].Visuals.OfType<HtmlRenderShape>().Count(shape => shape.Source?.StartsWith("div#limited:box-shadow[", StringComparison.Ordinal) == true));
        Assert.Contains("layers=3;limit=2", diagnostic.Detail, StringComparison.Ordinal);
        Assert.Equal(2, options.Clone().MaxBoxShadowLayers);
        Assert.Contains(HtmlRenderDiagnosticCodes.BoxShadowLayerLimit, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.BoxShadowLayerLimit, out _));

        options.MaxBoxShadowLayers = 0;
        Assert.Throws<ArgumentOutOfRangeException>(() => HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options));
    }

    [Fact]
    public void HtmlShadows_UnsupportedFormsUseCatalogedDiagnostics() {
        const string html = "<div id='invalid-shadow' style='width:20px;height:20px;box-shadow:0 1px -2px black'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 30D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported);

        Assert.Equal("div#invalid-shadow", diagnostic.Source);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#invalid-shadow:box-shadow");
        Assert.Contains(HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.BoxShadowValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(box-shadow:0 1px 2px rgba(0,0,0,.2))"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(box-shadow:0 1px 2px 0 black)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(box-shadow:inset 0 1px 2px black)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(box-shadow:0 1px 2px 3px black)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(box-shadow:0 1px black, 0 2px blue)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(box-shadow:0 1px -2px black)"));
    }
}
