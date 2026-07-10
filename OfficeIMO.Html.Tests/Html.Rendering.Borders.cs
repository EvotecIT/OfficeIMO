using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlBorders_DashedBorderAndDottedOutlineFlowThroughAllBackends() {
        const string html = "<div id='styled-strokes' style='width:30px;height:16px;margin:5px;color:#ff0000;border:2px dashed currentColor;border-radius:4px;outline:2px dotted #0000ff;outline-offset:3px;background:#ffffff;font-size:6px;line-height:8px'>StrokePdf</div>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 50D,
            ViewportHeight = 35D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderShape border = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#styled-strokes" && shape.Shape.StrokeWidth > 0D);
        HtmlRenderShape outline = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#styled-strokes:outline");
        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), visual => visual.Text == "StrokePdf");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(50D / HtmlRenderOptions.CssPixelsPerInch, 35D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Load(html.SaveAsPdf(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(OfficeStrokeDashStyle.Dash, border.Shape.StrokeDashStyle);
        Assert.Equal(OfficeColor.Red, border.Shape.StrokeColor);
        Assert.Equal(OfficeStrokeDashStyle.Dot, outline.Shape.StrokeDashStyle);
        Assert.Equal(OfficeColor.Blue, outline.Shape.StrokeColor);
        Assert.False(outline.Shape.Transform.HasValue);
        Assert.Equal(border.X - 4D, outline.X, 3);
        Assert.Equal(border.Width + 8D, outline.Width, 3);
        Assert.True(outline.PaintOrder > text.PaintOrder);
        Assert.Contains(
            Enumerable.Range(0, raster.Width).SelectMany(x => Enumerable.Range(0, raster.Height).Select(y => raster.GetPixel(x, y))),
            pixel => pixel.A > 0 && pixel.B > pixel.R);
        Assert.Contains("stroke-dasharray=", svg, StringComparison.Ordinal);
        Assert.Contains("StrokePdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlBorders_DoublePaintUsesTwoSharedVectorStrokes() {
        const string html = "<div id='double-border' style='width:30px;height:20px;margin:0;border:6px double #008000;border-radius:6px'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 35D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });
        List<HtmlRenderShape> borders = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source != null && shape.Source.StartsWith("div#double-border:border-", StringComparison.Ordinal))
            .ToList();

        Assert.Equal(2, borders.Count);
        Assert.All(borders, border => {
            Assert.Equal(2D, border.Shape.StrokeWidth, 3);
            Assert.Equal(OfficeColor.Green, border.Shape.StrokeColor);
            Assert.Equal(OfficeStrokeDashStyle.Solid, border.Shape.StrokeDashStyle);
        });
        HtmlRenderShape inner = Assert.Single(borders, border => border.Source == "div#double-border:border-inner");
        Assert.False(inner.Shape.Transform.HasValue);
        Assert.Equal(4D, inner.X, 3);
    }

    [Fact]
    public void HtmlBorders_NoneDoesNotConsumeLayoutInsetOrPaint() {
        const string html = "<div id='none-border' style='width:30px;height:10px;margin:0;border:8px none red;background:#00ff00'></div>"
            + "<div id='next-box' style='width:30px;height:10px;margin:0;background:#0000ff'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });
        HtmlRenderShape first = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#none-border");
        HtmlRenderShape next = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#next-box");

        Assert.Equal(30D, first.Width, 3);
        Assert.Equal(10D, next.Y, 3);
        Assert.Null(first.Shape.StrokeColor);
        Assert.Equal(0D, first.Shape.StrokeWidth, 3);
    }

    [Fact]
    public void HtmlBorders_ExplicitNoneSuppressesSyntheticTableCellBorder() {
        const string html = "<table style='width:30px;border:2px solid red;border-collapse:separate'><tr>"
            + "<td id='plain-cell' style='height:10px'></td><td id='none-cell' style='height:10px;border:none'></td>"
            + "</tr></table>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#plain-cell" && shape.Shape.StrokeWidth > 0D);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#none-cell" && shape.Shape.StrokeWidth > 0D);
    }

    [Fact]
    public void HtmlBorders_UnsupportedPaintUsesCatalogedDiagnosticsAndSupportsTruth() {
        const string html = "<div id='side-border' style='width:20px;height:10px;border-left:3px solid red'></div>"
            + "<div id='overridden-border' style='width:20px;height:10px;border:2px solid red;border-left-color:blue'></div>"
            + "<div id='groove-outline' style='width:20px;height:10px;outline:2px groove blue'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });
        List<HtmlDiagnostic> borderDiagnostics = rendered.Diagnostics.Diagnostics.Where(item => item.Code == HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported).ToList();
        HtmlDiagnostic outlineDiagnostic = Assert.Single(rendered.Diagnostics.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported);

        Assert.Equal(2, borderDiagnostics.Count);
        Assert.All(borderDiagnostics, diagnostic => Assert.Contains("asymmetric-side", diagnostic.Detail, StringComparison.Ordinal));
        Assert.Contains(borderDiagnostics, diagnostic => diagnostic.Source == "div#side-border");
        Assert.Contains(borderDiagnostics, diagnostic => diagnostic.Source == "div#overridden-border");
        Assert.Equal("div#groove-outline", outlineDiagnostic.Source);
        Assert.Contains("outline=", outlineDiagnostic.Detail, StringComparison.Ordinal);
        Assert.Contains(HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.Contains(HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported, out _));
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border:2px dashed red)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(outline:thin dotted currentColor)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-width:2px 2px 2px 2px)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(outline-offset:-2px)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border:2px groove red)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border-width:2px 3px)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border-left:2px dashed red)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(outline-offset:20%)"));
    }

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
