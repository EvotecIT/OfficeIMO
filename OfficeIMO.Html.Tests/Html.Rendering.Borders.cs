using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlBorders_BoxDecorationBreakPaintsInlineLineFragments() {
        const string html = "<p style='margin:0;width:72px;font-size:10px;line-height:12px'>"
            + "<span id='slice' style='background:#ffeecc;border:1px solid red;padding:1px;box-decoration-break:slice'>alpha beta gamma delta epsilon</span></p>"
            + "<p style='margin:0;width:72px;font-size:10px;line-height:12px'>"
            + "<span id='clone' style='background:#ccffee;border:1px solid blue;padding:1px;box-decoration-break:clone'>alpha beta gamma delta epsilon</span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        List<HtmlRenderShape> sliceBackgrounds = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == "span#slice" && shape.Shape.FillColor == OfficeColor.FromRgb(0xFF, 0xEE, 0xCC))
            .ToList();
        List<HtmlRenderShape> cloneBackgrounds = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == "span#clone" && shape.Shape.FillColor == OfficeColor.FromRgb(0xCC, 0xFF, 0xEE))
            .ToList();
        List<HtmlRenderShape> sliceEdges = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source != null && shape.Source.StartsWith("span#slice:border-", StringComparison.Ordinal))
            .ToList();
        List<HtmlRenderShape> cloneBorders = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == "span#clone" && shape.Shape.StrokeColor == OfficeColor.Blue)
            .ToList();

        Assert.True(sliceBackgrounds.Count >= 2);
        Assert.Equal(sliceBackgrounds.Count, cloneBackgrounds.Count);
        Assert.Single(sliceEdges, shape => shape.Source == "span#slice:border-left");
        Assert.Single(sliceEdges, shape => shape.Source == "span#slice:border-right");
        Assert.Equal(cloneBackgrounds.Count, cloneBorders.Count);
        Assert.All(cloneBorders, shape => Assert.Equal(OfficeColor.Blue, shape.Shape.StrokeColor));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(box-decoration-break:slice)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(box-decoration-break:clone)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(box-decoration-break:repeat)"));
    }

    [Fact]
    public void HtmlBorders_BoxDecorationBreakSliceRetainsOnlyDocumentOuterEdgesAcrossPages() {
        string words = string.Join(" ", Enumerable.Range(0, 36).Select(index => "word" + index.ToString("D2")));
        string html = "<style>@page{size:100px 40px;margin:0}html,body,p{margin:0}p{font-size:10px;line-height:12px;orphans:1;widows:1}</style>"
            + "<p><span id='paged-slice' style='background:#ffeecc;border:1px solid red;box-decoration-break:slice'>"
            + words
            + "</span></p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });
        List<HtmlRenderShape> edges = rendered.Pages
            .SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderShape>()
            .Where(shape => shape.Source != null && shape.Source.StartsWith("span#paged-slice:border-", StringComparison.Ordinal))
            .ToList();

        Assert.True(rendered.Pages.Count >= 2);
        Assert.Single(edges, shape => shape.Source == "span#paged-slice:border-left");
        Assert.Single(edges, shape => shape.Source == "span#paged-slice:border-right");
        Assert.DoesNotContain(rendered.Pages.Skip(1).SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderShape>(),
            shape => shape.Source == "span#paged-slice:border-left");
    }

    [Fact]
    public void HtmlBorders_InvertOutlineUsesDeterministicBackdropInversion() {
        const string html = "<div id='invert' style='width:24px;height:12px;background:#123456;outline:2px solid invert'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape outline = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#invert:outline");
        Assert.Equal(OfficeColor.FromRgb(0xED, 0xCB, 0xA9), outline.Shape.StrokeColor);
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(outline-color:invert)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(outline:2px solid invert)"));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported);
    }

    [Fact]
    public void HtmlBorders_DashedBorderAndDottedOutlineFlowThroughAllBackends() {
        const string html = "<div id='styled-strokes' style='width:30px;height:16px;margin:5px;color:#ff0000;border:2px dashed currentColor;border-radius:4px;outline:2px dotted #0000ff;outline-offset:3px;background:#ffffff;font-size:6px;line-height:8px'>StrokePdf</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 35D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape border = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#styled-strokes" && shape.Shape.StrokeWidth > 0D);
        HtmlRenderShape outline = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#styled-strokes:outline");
        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), visual => visual.Text == "StrokePdf");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlToPdfOptions pdfOptions = new HtmlToPdfOptions();
        pdfOptions = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(50D / HtmlRenderOptions.CssPixelsPerInch, 35D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfBytes(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

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
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlBorders_DoublePaintUsesTwoSharedVectorStrokes() {
        const string html = "<div id='double-border' style='width:30px;height:20px;margin:0;border:6px double #008000;border-radius:6px'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#plain-cell" && shape.Shape.StrokeWidth > 0D);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#none-cell" && shape.Shape.StrokeWidth > 0D);
    }

    [Fact]
    public void HtmlBorders_SideSpecificPaintUsesSharedGeometryAndTruthfulSupports() {
        const string html = "<div id='side-border' style='width:20px;height:10px;border-left:3px solid red;background:#ffffff'></div>"
            + "<div id='overridden-border' style='width:20px;height:10px;border:2px solid red;border-left-color:blue'></div>"
            + "<div id='mixed-border' style='width:32px;height:18px;border-width:2px 3px 4px 5px;border-style:solid dashed dotted double;border-color:red green blue black;border-radius:10px 3px / 4px 8px;background:#ffffff;font-size:6px;line-height:8px'>BorderPdf</div>"
            + "<div id='groove-border' style='width:20px;height:10px;border-left:2px groove red'></div>"
            + "<div id='groove-outline' style='width:20px;height:10px;outline:2px groove blue'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 70D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        List<HtmlRenderShape> mixed = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source != null && shape.Source.StartsWith("div#mixed-border:border-", StringComparison.Ordinal))
            .ToList();
        HtmlRenderShape sideBackground = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#side-border");
        HtmlRenderShape mixedBackground = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#mixed-border" && shape.Shape.FillColor == OfficeColor.White);
        HtmlRenderText mixedText = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "BorderPdf");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlToPdfOptions pdfOptions = new HtmlToPdfOptions();
        pdfOptions = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(70D / HtmlRenderOptions.CssPixelsPerInch, 80D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfBytes(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(23D, sideBackground.Width, 3);
        Assert.Equal(40D, mixedBackground.Width, 3);
        Assert.Equal(24D, mixedBackground.Height, 3);
        Assert.Equal(5D, mixedText.X, 3);
        Assert.Equal(5, mixed.Count);
        Assert.Contains(mixed, shape => shape.Source == "div#mixed-border:border-top" && shape.Shape.StrokeWidth == 2D && shape.Shape.StrokeColor == OfficeColor.Red);
        Assert.Contains(mixed, shape => shape.Source == "div#mixed-border:border-right" && shape.Shape.StrokeWidth == 3D && shape.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dash && shape.Shape.StrokeColor == OfficeColor.Green);
        Assert.Contains(mixed, shape => shape.Source == "div#mixed-border:border-bottom" && shape.Shape.StrokeWidth == 4D && shape.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dot && shape.Shape.StrokeColor == OfficeColor.Blue);
        Assert.Equal(2, mixed.Count(shape => shape.Source != null && shape.Source.StartsWith("div#mixed-border:border-left-", StringComparison.Ordinal) && Math.Abs(shape.Shape.StrokeWidth - 5D / 3D) < 0.0001D));
        Assert.All(mixed, shape => Assert.Equal(OfficeShapeKind.Path, shape.Shape.Kind));
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape =>
            shape.Source != null && shape.Source.StartsWith("div#groove-border:border-left", StringComparison.Ordinal));
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape =>
            shape.Source != null && shape.Source.StartsWith("div#groove-outline:outline:border-", StringComparison.Ordinal));
        Assert.Contains(Enumerable.Range(0, raster.Width).SelectMany(x => Enumerable.Range(0, raster.Height).Select(y => raster.GetPixel(x, y))), pixel => pixel.A > 0 && pixel.B > pixel.R);
        Assert.Contains("<path", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-dasharray=", svg, StringComparison.Ordinal);
        Assert.Contains("BorderPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Source == "div#side-border" || diagnostic.Source == "div#overridden-border" || diagnostic.Source == "div#mixed-border");
        Assert.Contains(HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.Contains(HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported, out _));
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border:2px dashed red)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(outline:thin dotted currentColor)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-width:2px 3px 4px 5px)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-style:solid dashed dotted double)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-color:red green blue black)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-left:2px dashed red)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-right-width:3px)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(outline-offset:-2px)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border:2px groove red)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-left:2px groove red)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(outline:2px ridge red)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(outline-offset:20%)"));
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlBorders_ThreeDimensionalStylesUseSharedTwoToneVectorPaintAcrossBackends() {
        const string html = "<div id='groove' style='width:24px;height:12px;border:6px groove #808080'>Groove</div>"
            + "<div id='ridge' style='width:24px;height:12px;border:6px ridge #808080'>Ridge</div>"
            + "<div id='inset' style='width:24px;height:12px;border:6px inset #808080'>Inset</div>"
            + "<div id='outset' style='width:24px;height:12px;border:6px outset #808080'>OutsetPdf</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 60D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(2D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        List<HtmlRenderShape> groove = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source != null && shape.Source.StartsWith("div#groove:border-", StringComparison.Ordinal))
            .ToList();
        List<HtmlRenderShape> ridge = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source != null && shape.Source.StartsWith("div#ridge:border-", StringComparison.Ordinal))
            .ToList();
        List<HtmlRenderShape> inset = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source != null && shape.Source.StartsWith("div#inset:border-", StringComparison.Ordinal))
            .ToList();
        List<HtmlRenderShape> outset = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source != null && shape.Source.StartsWith("div#outset:border-", StringComparison.Ordinal))
            .ToList();
        OfficeColor dark = OfficeColorTransforms.Shade(OfficeColor.Gray, 0.55D);
        OfficeColor light = OfficeColorTransforms.Tint(OfficeColor.Gray, 0.55D);
        string svg = HtmlConversionDocument.Parse(html).ToSvg(options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(options));

        Assert.Equal(8, groove.Count);
        Assert.Equal(8, ridge.Count);
        Assert.Equal(4, inset.Count);
        Assert.Equal(4, outset.Count);
        Assert.All(new[] { groove, ridge, inset, outset }, shapes => {
            Assert.Contains(shapes, shape => shape.Shape.StrokeColor == dark);
            Assert.Contains(shapes, shape => shape.Shape.StrokeColor == light);
            Assert.All(shapes, shape => Assert.Equal(OfficeStrokeDashStyle.Solid, shape.Shape.StrokeDashStyle));
        });
        Assert.Contains("<path", svg, StringComparison.Ordinal);
        Assert.Contains("OutsetPdf", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported);
    }

    [Fact]
    public void HtmlBorders_UniformRadiusFlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<div id='rounded' style='width:40px;height:20px;margin:0;border:2px solid #0000ff;border-radius:6px;background:#ff0000;font-size:6px;line-height:8px'>RoundedPdf</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        List<HtmlRenderShape> shapes = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source == "div#rounded")
            .ToList();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlToPdfOptions pdfOptions = new HtmlToPdfOptions();
        pdfOptions = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(50D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfBytes(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(2, shapes.Count);
        Assert.All(shapes, shape => {
            Assert.Equal(OfficeShapeKind.RoundedRectangle, shape.Shape.Kind);
            Assert.Equal(6D, shape.Shape.CornerRadius, 3);
        });
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(0, 0));
        Assert.Equal((byte)255, raster.GetPixel(20, 15).R);
        Assert.Contains("rx=\"6\" ry=\"6\"", svg, StringComparison.Ordinal);
        Assert.Contains("RoundedPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlBorders_AsymmetricEllipticalRadiusFlowsThroughSharedPathBackends() {
        const string html = "<div id='asymmetric' style='width:40px;height:24px;margin:3px;border:2px dashed #0000ff;outline:2px dotted #008000;outline-offset:1px;border-radius:12px 3px 8px 5px / 4px 9px 2px 7px;background-color:#ffffff;background-image:linear-gradient(90deg,#ff0000,#00ff00);box-shadow:2px 1px 0 #000000;font-size:6px;line-height:8px'>PathPdf</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 50D,
            ViewportHeight = 35D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        List<HtmlRenderShape> shapes = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(item => item.Shape.Kind == OfficeShapeKind.Path && item.Source != null
                && (item.Source == "div#asymmetric" || item.Source.StartsWith("div#asymmetric:outline", StringComparison.Ordinal)))
            .ToList();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlToPdfOptions pdfOptions = new HtmlToPdfOptions();
        pdfOptions = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(50D / HtmlRenderOptions.CssPixelsPerInch, 35D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfBytes(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.True(shapes.Count >= 2);
        Assert.All(shapes, shape => Assert.Equal(OfficeShapeKind.Path, shape.Shape.Kind));
        Assert.All(shapes, shape => Assert.True(shape.Shape.PathCommands.Count >= 10));
        Assert.True(raster.GetPixel(20, 15).A > 0);
        Assert.Contains("<path", svg, StringComparison.Ordinal);
        Assert.Contains("PathPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-radius:6px 2px)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-radius:12px / 4px)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-top-left-radius:6px 2px)"));
    }

    [Fact]
    public void HtmlBorders_CalculatedRadiusUsesRoundedPath() {
        const string html = "<div id='invalid-radius' style='width:30px;height:20px;margin:0;border-radius:calc(6px);background:#ff0000'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });
        HtmlRenderShape shape = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), item => item.Source == "div#invalid-radius");
        Assert.Equal(OfficeShapeKind.RoundedRectangle, shape.Shape.Kind);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported);
        Assert.Contains(HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.BorderRadiusValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-radius:6px)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-top-left-radius:6px)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-radius:calc(6px))"));
    }

    [Fact]
    public void HtmlBorders_ContainerUnitsUseTheActiveQueryContainerForEveryRadiusAxis() {
        const string html = "<section style='width:200px;height:80px;container-type:size'><div id='container-radius' style='width:100px;height:60px;border-radius:10cqw / 5cqh;background:red'></div></section>";

        HtmlRenderShape shape = Assert.Single(HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 1000D,
            ViewportHeight = 800D,
            Margins = HtmlRenderMargins.All(0D)
        }).Pages[0].Visuals.OfType<HtmlRenderShape>(), item => item.Source == "div#container-radius");

        Assert.Equal(OfficeShapeKind.Path, shape.Shape.Kind);
        Assert.Equal(20D, shape.Shape.PathCommands[0].Point.X, 3);
        Assert.Equal(4D, shape.Shape.PathCommands[2].Point.Y, 3);
    }

    [Fact]
    public void HtmlBorders_RadiusOverlapUsesCssProportionalNormalization() {
        const string html = "<div id='normalized-radius' style='width:40px;height:20px;margin:0;border-radius:30px 20px 10px 5px / 20px 20px 10px 5px;background:#ff0000'></div>";

        HtmlRenderShape shape = Assert.Single(HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 45D,
            ViewportHeight = 25D,
            Margins = HtmlRenderMargins.All(0D)
        }).Pages[0].Visuals.OfType<HtmlRenderShape>(), item => item.Source == "div#normalized-radius");

        Assert.Equal(OfficeShapeKind.Path, shape.Shape.Kind);
        Assert.Equal(20D, shape.Shape.PathCommands[0].Point.X, 3);
        Assert.Equal(40D - (20D * 2D / 3D), shape.Shape.PathCommands[1].Point.X, 3);
    }

    [Fact]
    public void HtmlBorders_OversizedAsymmetricPathContinuesThroughClippedPageFragments() {
        const string html = "<html style='margin:0'><body style='margin:0'><div id='tall-path' style='box-sizing:border-box;width:40px;height:90px;border-radius:12px 3px 8px 5px / 4px 9px 2px 7px;background:#ff0000'></div></body></html>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(40D / HtmlRenderOptions.CssPixelsPerInch, 40D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        IReadOnlyList<OfficeImageExportResult> pngPages = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Png, options);
        IReadOnlyList<OfficeImageExportResult> svgPages = HtmlConversionDocument.Parse(html).ExportImages(OfficeImageExportFormat.Svg, options);
        HtmlToPdfOptions pdfOptions = new HtmlToPdfOptions();
        pdfOptions = new HtmlToPdfOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfBytes(pdfOptions);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => {
            HtmlRenderClipGroup fragment = Assert.Single(page.Visuals.OfType<HtmlRenderClipGroup>(), group =>
                group.Source == "div#tall-path"
                && group.Visuals.OfType<HtmlRenderShape>().Any(shape => shape.Shape.Kind == OfficeShapeKind.Path));
            Assert.Single(fragment.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.Kind == OfficeShapeKind.Path);
        });
        for (int index = 0; index < 3; index++) {
            Assert.True(OfficePngReader.TryDecode(pngPages[index].Bytes, out OfficeRasterImage? raster));
            Assert.True(raster!.GetPixel(20, index < 2 ? 20 : 5).R > 240);
            string svg = Encoding.UTF8.GetString(svgPages[index].Bytes);
            Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
            Assert.Contains("<path", svg, StringComparison.Ordinal);
        }
        Assert.Equal(3, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }
}
