using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlLinearGradient_FlowsAsMultiStopVectorPaintAcrossPngSvgAndSearchablePdf() {
        const string html = "<div style=\"width:160px;height:30px;background-image:linear-gradient(to right,#ff0000 0%,#00ff00 50%,#0000ff 100%)\">GradientMarker</div>";
        var imageOptions = new HtmlImageExportOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, imageOptions);
        HtmlRenderShape gradientShape = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null);
        OfficeLinearGradient gradient = gradientShape.Shape.FillGradient!;
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = html.ToSvg(imageOptions);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        byte[] pdf = html.SaveAsPdf(pdfOptions);

        Assert.Equal(3, gradient.Stops.Count);
        Assert.Equal(0D, gradient.StartX, 3);
        Assert.Equal(1D, gradient.EndX, 3);
        Assert.True(raster.GetPixel(9, 20).R > raster.GetPixel(9, 20).G);
        Assert.True(raster.GetPixel(88, 20).G > raster.GetPixel(88, 20).R);
        Assert.True(raster.GetPixel(166, 20).B > raster.GetPixel(166, 20).G);
        Assert.Contains("<linearGradient", svg, StringComparison.Ordinal);
        Assert.Equal(3, CountBackgroundOccurrences(svg, "<stop "));
        Assert.Contains("/FunctionType 3", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        Assert.Contains("GradientMarker", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlRadialGradient_FlowsAsMultiStopVectorPaintAcrossPngSvgAndSearchablePdf() {
        const string html = "<div style=\"width:160px;height:60px;background-image:radial-gradient(ellipse at center,#ff0000 0%,#00ff00 50%,#0000ff 100%)\">RadialMarker</div>";
        var imageOptions = new HtmlImageExportOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, imageOptions);
        HtmlRenderShape gradientShape = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null);
        OfficeRadialGradient gradient = gradientShape.Shape.FillRadialGradient!;
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = html.ToSvg(imageOptions);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = imageOptions;
        byte[] pdf = html.SaveAsPdf(pdfOptions);
        string pdfSource = Encoding.ASCII.GetString(pdf);

        Assert.Equal(3, gradient.Stops.Count);
        Assert.Equal(Math.Sqrt(0.5D), gradient.EndRadius, 3);
        Assert.True(raster.GetPixel(88, 38).R > raster.GetPixel(88, 38).B);
        Assert.True(raster.GetPixel(9, 66).B > raster.GetPixel(9, 66).R);
        Assert.Contains("<radialGradient", svg, StringComparison.Ordinal);
        Assert.Equal(3, CountBackgroundOccurrences(svg, "<stop "));
        Assert.Contains("/ShadingType 3", pdfSource, StringComparison.Ordinal);
        Assert.Contains("/Coords [0.5 0.5 0 0.5 0.5 0.707]", pdfSource, StringComparison.Ordinal);
        Assert.Contains("RadialMarker", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlRadialGradient_OffCenterEllipseFlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<div style='width:160px;height:60px;background:radial-gradient(ellipse farthest-side at 25% 50%,red,blue)'></div><p>EllipseMarker</p>";
        var imageOptions = new HtmlImageExportOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, imageOptions);
        OfficeRadialGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null).Shape.FillRadialGradient!;
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = html.ToSvg(imageOptions);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = imageOptions;
        byte[] pdf = html.SaveAsPdf(pdfOptions);
        string pdfSource = Encoding.ASCII.GetString(pdf);

        Assert.Equal(0.25D, gradient.EndX, 3);
        Assert.Equal(0.5D, gradient.EndY, 3);
        Assert.Equal(0.75D, gradient.EndRadiusX, 3);
        Assert.Equal(0.5D, gradient.EndRadiusY, 3);
        Assert.True(raster.GetPixel(48, 38).R > raster.GetPixel(48, 38).B);
        Assert.True(raster.GetPixel(167, 38).B > raster.GetPixel(167, 38).R);
        Assert.Contains("gradientTransform=\"matrix(0.75 0 0 0.5 0.25 0.5)\"", svg, StringComparison.Ordinal);
        Assert.Contains("/ShadingType 3", pdfSource, StringComparison.Ordinal);
        Assert.Contains("/Coords [0 0 0 0 0 1]", pdfSource, StringComparison.Ordinal);
        Assert.Contains("EllipseMarker", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlRadialGradient_DiagnosesCircularFormsPendingPaintAreaResolution() {
        const string html = "<div style='width:40px;height:20px;background:radial-gradient(circle at left,red,blue)'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        });

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.FillRadialGradient != null);
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Theory]
    [InlineData("ellipse closest-side at center", 0.5D)]
    [InlineData("farthest-corner ellipse at 50% 50%", 0.7071067811865476D)]
    public void HtmlRadialGradient_MapsCenteredEllipseSizeKeywords(string descriptor, double radius) {
        string html = "<div style='width:40px;height:20px;background:radial-gradient(" + descriptor + ",red,blue)'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        });

        OfficeRadialGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null).Shape.FillRadialGradient!;
        Assert.Equal(radius, gradient.EndRadius, 3);
    }

    [Theory]
    [InlineData("ellipse farthest-side at 25% 50%", 0.25D, 0.5D, 0.75D, 0.5D)]
    [InlineData("farthest-corner ellipse at right top", 1D, 0D, 1.4142135623730951D, 1.4142135623730951D)]
    [InlineData("ellipse 20% 35% at 30% 70%", 0.3D, 0.7D, 0.2D, 0.35D)]
    [InlineData("ellipse farthest-side at -25% 150%", -0.25D, 1.5D, 1.25D, 1.5D)]
    public void HtmlRadialGradient_MapsEllipsePositionAndIndependentRadii(
        string descriptor,
        double centerX,
        double centerY,
        double radiusX,
        double radiusY) {
        string html = "<div style='width:40px;height:20px;background:radial-gradient(" + descriptor + ",red,blue)'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        });

        OfficeRadialGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null).Shape.FillRadialGradient!;
        Assert.Equal(centerX, gradient.EndX, 3);
        Assert.Equal(centerY, gradient.EndY, 3);
        Assert.Equal(radiusX, gradient.EndRadiusX, 3);
        Assert.Equal(radiusY, gradient.EndRadiusY, 3);
    }

    [Theory]
    [InlineData("to right", 0D, 0.5D, 1D, 0.5D)]
    [InlineData("to bottom", 0.5D, 0D, 0.5D, 1D)]
    [InlineData("to top left", 1D, 1D, 0D, 0D)]
    [InlineData("90deg", 0D, 0.5D, 1D, 0.5D)]
    public void HtmlLinearGradient_MapsCssDirectionsToDrawingCoordinates(
        string direction,
        double startX,
        double startY,
        double endX,
        double endY) {
        string html = "<div style='width:40px;height:20px;background:linear-gradient(" + direction + ",red,blue)'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        });

        OfficeLinearGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        Assert.Equal(startX, gradient.StartX, 3);
        Assert.Equal(startY, gradient.StartY, 3);
        Assert.Equal(endX, gradient.EndX, 3);
        Assert.Equal(endY, gradient.EndY, 3);
    }

    [Fact]
    public void HtmlLinearGradient_DistributesImplicitStopsAndExtendsEndpointColors() {
        const string html = "<div style='width:40px;height:20px;background:linear-gradient(red 20%,lime,blue 80%)'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        });

        OfficeLinearGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        Assert.Equal(new[] { 0D, 0.2D, 0.5D, 0.8D, 1D }, gradient.Stops.Select(stop => stop.Offset));
        Assert.Equal(OfficeColor.Red, gradient.Stops[0].Color);
        Assert.Equal(OfficeColor.Red, gradient.Stops[1].Color);
        Assert.Equal(OfficeColor.Blue, gradient.Stops[3].Color);
        Assert.Equal(OfficeColor.Blue, gradient.Stops[4].Color);
    }

    [Fact]
    public void HtmlLinearGradient_ComposesAboveUrlLayers() {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(0, 0, 255));
        string html = "<div style=\"width:40px;height:20px;background-image:linear-gradient(to right,red,lime),url('data:image/png;base64,"
            + imageData
            + "');background-size:auto,40px 20px;background-repeat:no-repeat\"></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        });

        IReadOnlyList<HtmlRenderVisual> backgrounds = rendered.Pages[0].Visuals
            .Where(visual => visual.Source != null && visual.Source.IndexOf(":background-", StringComparison.Ordinal) >= 0)
            .ToList();
        Assert.Equal(2, backgrounds.Count);
        Assert.IsType<HtmlRenderImage>(backgrounds[0]);
        Assert.EndsWith(":background-image[1]", backgrounds[0].Source, StringComparison.Ordinal);
        Assert.IsType<HtmlRenderShape>(backgrounds[1]);
        Assert.EndsWith(":background-gradient[0]", backgrounds[1].Source, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlLinearGradient_PropagatesFromTheRootAcrossTheWholeCanvas() {
        const string html = "<style>body{background:linear-gradient(to bottom,red,blue)}</style><p>G</p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderShape gradient = Assert.Single(
            page.Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "render-root-background:background-gradient");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());

        Assert.Equal(page.Width, gradient.Width, 3);
        Assert.Equal(page.Height, gradient.Height, 3);
        Assert.True(raster.GetPixel(raster.Width / 2, 1).R > raster.GetPixel(raster.Width / 2, 1).B);
        Assert.True(raster.GetPixel(raster.Width / 2, raster.Height - 2).B > raster.GetPixel(raster.Width / 2, raster.Height - 2).R);
    }

    [Theory]
    [InlineData("linear-gradient(red,lime,blue)")]
    [InlineData("radial-gradient(red,lime,blue)")]
    public void HtmlRender_BoundsCssGradientStops(string background) {
        string html = "<div style='width:40px;height:20px;background:" + background + "'></div>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D),
            MaxGradientStops = 2
        });

        Assert.DoesNotContain(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null || shape.Shape.FillRadialGradient != null);
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GradientStopLimitExceeded);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }
}
