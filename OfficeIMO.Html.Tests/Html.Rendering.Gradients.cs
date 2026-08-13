using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlLinearGradient_PremultipliesSynthesizedBoundaryStops() {
        const string html = "<div style='width:100px;height:20px;background:linear-gradient(to right,rgba(255 0 0 / 0) -100%,blue 100%)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        OfficeLinearGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;

        Assert.Equal(0D, gradient.Stops[0].Offset, 3);
        Assert.Equal(OfficeColor.FromRgba(0, 0, 255, 128), gradient.Stops[0].Color);
    }

    [Fact]
    public void HtmlLinearGradient_FlowsAsMultiStopVectorPaintAcrossPngSvgAndSearchablePdf() {
        const string html = "<div style=\"width:160px;height:30px;background-image:linear-gradient(to right,#ff0000 0%,#00ff00 50%,#0000ff 100%)\">GradientMarker</div>";
        var imageOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), imageOptions);
        HtmlRenderShape gradientShape = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null);
        OfficeLinearGradient gradient = gradientShape.Shape.FillGradient!;
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = HtmlConversionDocument.Parse(html).ToSvg(imageOptions);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);

        Assert.Equal(3, gradient.Stops.Count);
        Assert.Equal(0D, gradient.StartX, 3);
        Assert.Equal(1D, gradient.EndX, 3);
        const int unobscuredY = 34;
        Assert.True(raster.GetPixel(9, unobscuredY).R > raster.GetPixel(9, unobscuredY).G);
        Assert.True(raster.GetPixel(88, unobscuredY).G > raster.GetPixel(88, unobscuredY).R);
        Assert.True(raster.GetPixel(166, unobscuredY).B > raster.GetPixel(166, unobscuredY).G);
        Assert.Contains("<linearGradient", svg, StringComparison.Ordinal);
        Assert.Equal(3, CountBackgroundOccurrences(svg, "<stop "));
        Assert.Contains("/FunctionType 3", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        Assert.Contains("GradientMarker", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlLinearGradient_PreservesHardStopsAcrossPngSvgAndSearchablePdf() {
        const string html = "<div style='width:100px;height:20px;background:linear-gradient(to right,red 0%,red 50%,blue 50%,blue 100%)'></div><p>HardStopPdf</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 130D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeLinearGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = HtmlConversionDocument.Parse(html).ToSvg(options);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfSource = Encoding.ASCII.GetString(pdf);

        Assert.Equal(new[] { 0D, 0.5D, 0.5D, 1D }, gradient.Stops.Select(stop => stop.Offset));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(55, 15));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(61, 15));
        Assert.Equal(2, CountBackgroundOccurrences(svg, "offset=\"50%\""));
        Assert.Contains("/Bounds [0.4999999 0.5000001]", pdfSource, StringComparison.Ordinal);
        Assert.Contains("HardStopPdf", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlLinearGradient_ResolvesLengthStopsAgainstThePhysicalGradientLine() {
        const string html = "<div style='width:100px;height:20px;background:linear-gradient(to right,red 0px,lime 25px,blue 100px)'></div><p>LengthStopPdf</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 130D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeLinearGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        string svg = HtmlConversionDocument.Parse(html).ToSvg(options);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);

        Assert.Equal(new[] { 0D, 0.25D, 1D }, gradient.Stops.Select(stop => stop.Offset));
        Assert.Contains("offset=\"25%\"", svg, StringComparison.Ordinal);
        Assert.Contains("/Bounds [0.25]", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        Assert.Contains("LengthStopPdf", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlRadialGradient_ResolvesLengthStopsAgainstTheResolvedGradientRadius() {
        const string html = "<div style='width:100px;height:60px;background:radial-gradient(40px circle at center,red 0,lime 20px,blue 40px)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 130D,
            Margins = HtmlRenderMargins.All(8D)
        });

        OfficeRadialGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null).Shape.FillRadialGradient!;
        Assert.Equal(new[] { 0D, 0.5D, 1D }, gradient.Stops.Select(stop => stop.Offset));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlRadialGradient_FlowsAsMultiStopVectorPaintAcrossPngSvgAndSearchablePdf() {
        const string html = "<div style=\"width:160px;height:60px;background-image:radial-gradient(ellipse at center,#ff0000 0%,#00ff00 50%,#0000ff 100%)\">RadialMarker</div>";
        var imageOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), imageOptions);
        HtmlRenderShape gradientShape = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null);
        OfficeRadialGradient gradient = gradientShape.Shape.FillRadialGradient!;
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = HtmlConversionDocument.Parse(html).ToSvg(imageOptions);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(imageOptions);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfSource = Encoding.ASCII.GetString(pdf);

        Assert.Equal(3, gradient.Stops.Count);
        Assert.Equal(Math.Sqrt(0.5D), gradient.EndRadius, 3);
        Assert.True(raster.GetPixel(88, 38).R > raster.GetPixel(88, 38).B);
        Assert.True(raster.GetPixel(9, 66).B > raster.GetPixel(9, 66).R);
        Assert.Contains("<radialGradient", svg, StringComparison.Ordinal);
        Assert.Equal(3, CountBackgroundOccurrences(svg, "<stop "));
        Assert.Contains("/ShadingType 3", pdfSource, StringComparison.Ordinal);
        Assert.Contains("/Coords [0.5 0.5 0 0.5 0.5 0.707]", pdfSource, StringComparison.Ordinal);
        Assert.Contains("RadialMarker", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlRadialGradient_OffCenterEllipseFlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<div style='width:160px;height:60px;background:radial-gradient(ellipse farthest-side at 25% 50%,red,blue)'></div><p>EllipseMarker</p>";
        var imageOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), imageOptions);
        OfficeRadialGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null).Shape.FillRadialGradient!;
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = HtmlConversionDocument.Parse(html).ToSvg(imageOptions);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(imageOptions);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
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
        Assert.Contains("EllipseMarker", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlRadialGradient_CircleFlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<div style='width:160px;height:60px;background:radial-gradient(circle closest-side at 25% 50%,red,blue)'></div><p>CircleMarker</p>";
        var imageOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(8D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), imageOptions);
        OfficeRadialGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null).Shape.FillRadialGradient!;
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = HtmlConversionDocument.Parse(html).ToSvg(imageOptions);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(imageOptions);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);

        Assert.Equal(0.1875D, gradient.EndRadiusX, 4);
        Assert.Equal(0.5D, gradient.EndRadiusY, 4);
        Assert.True(raster.GetPixel(48, 38).R > raster.GetPixel(48, 38).B);
        Assert.True(raster.GetPixel(78, 38).B > raster.GetPixel(78, 38).R);
        Assert.Contains("gradientTransform=\"matrix(0.188 0 0 0.5 0.25 0.5)\"", svg, StringComparison.Ordinal);
        Assert.Contains("/ShadingType 3", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        Assert.Contains("CircleMarker", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlRadialGradient_ResolvesCircularExtentAgainstThePaintArea() {
        const string html = "<div style='width:40px;height:20px;background:radial-gradient(circle at left,red,blue)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        });

        OfficeRadialGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null).Shape.FillRadialGradient!;
        Assert.Equal(0D, gradient.EndX, 3);
        Assert.Equal(0.5D, gradient.EndY, 3);
        Assert.Equal(Math.Sqrt(1700D) / 40D, gradient.EndRadiusX, 3);
        Assert.Equal(Math.Sqrt(1700D) / 20D, gradient.EndRadiusY, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlRadialGradient_DegenerateCircleKeepsAUniformPhysicalRadius() {
        const string html = "<div style='width:40px;height:20px;background:radial-gradient(circle closest-side at left,red,blue)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        });

        OfficeRadialGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null).Shape.FillRadialGradient!;
        Assert.True(gradient.EndRadiusX > 0D);
        Assert.Equal(gradient.EndRadiusX * 40D, gradient.EndRadiusY * 20D, 9);
    }

    [Theory]
    [InlineData("8px circle at 10px 5px", 0.25D, 0.25D, 0.2D, 0.4D)]
    [InlineData("12px at center", 0.5D, 0.5D, 0.3D, 0.6D)]
    [InlineData("ellipse 25% 10px at 25% 75%", 0.25D, 0.75D, 0.25D, 0.5D)]
    public void HtmlRadialGradient_ResolvesExplicitLengthGeometry(
        string descriptor,
        double centerX,
        double centerY,
        double radiusX,
        double radiusY) {
        string html = "<div style='width:40px;height:20px;background:radial-gradient(" + descriptor + ",red,blue)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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
    [InlineData("circle 20%")]
    [InlineData("ellipse 12px")]
    [InlineData("circle 10px 20px")]
    [InlineData("ellipse -10px 20px")]
    [InlineData("circle at")]
    public void HtmlRadialGradient_DiagnosesInvalidShapeAndSizeCombinations(string descriptor) {
        string html = "<div style='width:40px;height:20px;background:radial-gradient(" + descriptor + ",red,blue)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        });

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.FillRadialGradient != null);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Theory]
    [InlineData("ellipse closest-side at center", 0.5D)]
    [InlineData("farthest-corner ellipse at 50% 50%", 0.7071067811865476D)]
    public void HtmlRadialGradient_MapsCenteredEllipseSizeKeywords(string descriptor, double radius) {
        string html = "<div style='width:40px;height:20px;background:radial-gradient(" + descriptor + ",red,blue)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

    [Theory]
    [InlineData("linear-gradient(135deg,red,blue)")]
    [InlineData("repeating-linear-gradient(135deg,red 0 10px,blue 10px 20px)")]
    public void HtmlLinearGradient_ExplicitAnglesPreservePhysicalDirectionOnNonSquareBoxes(string background) {
        string html = "<div style='width:200px;height:100px;background:" + background + "'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(0D)
        });

        OfficeLinearGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        double physicalX = (gradient.EndX - gradient.StartX) * 200D;
        double physicalY = (gradient.EndY - gradient.StartY) * 100D;
        Assert.Equal(physicalX, physicalY, 6);
        Assert.True(physicalX > 0D);
    }

    [Fact]
    public void HtmlLinearGradient_DirectionalCornerKeywordsKeepMagicCornerGeometry() {
        const string html = "<div style='width:200px;height:100px;background:linear-gradient(to bottom right,red,blue)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(0D)
        });

        OfficeLinearGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        Assert.Equal(0D, gradient.StartX, 3);
        Assert.Equal(0D, gradient.StartY, 3);
        Assert.Equal(1D, gradient.EndX, 3);
        Assert.Equal(1D, gradient.EndY, 3);
    }

    [Fact]
    public void HtmlLinearGradient_DistributesImplicitStopsAndExtendsEndpointColors() {
        const string html = "<div style='width:40px;height:20px;background:linear-gradient(red 20%,lime,blue 80%)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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
    public void HtmlLinearGradient_ClampsBackwardStopPositionsIntoAHardEdge() {
        const string html = "<div style='width:40px;height:20px;background:linear-gradient(red 60%,blue 40%,lime 100%)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D)
        });

        OfficeLinearGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        Assert.Equal(new[] { 0D, 0.6D, 0.6D, 1D }, gradient.Stops.Select(stop => stop.Offset));
        Assert.Equal(OfficeColor.Red, gradient.Stops[1].Color);
        Assert.Equal(OfficeColor.Blue, gradient.Stops[2].Color);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlLinearGradient_ExpandsTwoPositionColorStops() {
        const string html = "<div style='width:100px;height:20px;background:linear-gradient(to right,red 0 50%,blue 50% 100%)'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 130D,
            Margins = HtmlRenderMargins.All(8D)
        });

        OfficeLinearGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        Assert.Equal(new[] { 0D, 0.5D, 0.5D, 1D }, gradient.Stops.Select(stop => stop.Offset));
        Assert.Equal(new[] { OfficeColor.Red, OfficeColor.Red, OfficeColor.Blue, OfficeColor.Blue }, gradient.Stops.Select(stop => stop.Color));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlRepeatingLinearGradient_ExpandsItsAuthoredPeriodAcrossEveryManagedExporter() {
        const string html = "<div style='width:100px;height:20px;background:repeating-linear-gradient(to right,rgba(255 0 0 / 50%) 0 10px,color-mix(in srgb,blue 75%,white) 10px 20px)'></div><p>RepeatingLinear</p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 130D,
            Margins = HtmlRenderMargins.All(8D),
            MaxGradientStops = 32
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        OfficeLinearGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = HtmlConversionDocument.Parse(html).ToSvg(options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions(options));

        Assert.Equal(22, gradient.Stops.Count);
        Assert.Equal(0D, gradient.Stops[0].Offset, 6);
        Assert.Equal(1D, gradient.Stops[gradient.Stops.Count - 1].Offset, 6);
        OfficeGradientStop[] boundaryStops = gradient.Stops.Where(stop => stop.Offset == 0D).ToArray();
        Assert.Equal(2, boundaryStops.Length);
        Assert.Equal((byte)255, boundaryStops[0].Color.A);
        Assert.Equal((byte)128, boundaryStops[1].Color.A);
        Assert.NotEqual(raster.GetPixel(12, 15), raster.GetPixel(23, 15));
        Assert.Contains("<linearGradient", svg, StringComparison.Ordinal);
        Assert.Contains("RepeatingLinear", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlRepeatingRadialGradient_ExpandsItsPeriodAndHonorsTheMaterializedStopBudget() {
        const string background = "repeating-radial-gradient(circle,red 0 5px,blue 5px 10px)";
        string html = "<div style='width:40px;height:40px;background:" + background + "'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 70D,
            Margins = HtmlRenderMargins.All(8D),
            MaxGradientStops = 32
        });
        OfficeRadialGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null).Shape.FillRadialGradient!;
        Assert.True(gradient.Stops.Count > 4);
        Assert.Equal(0D, gradient.Stops[0].Offset, 6);
        Assert.Equal(1D, gradient.Stops[gradient.Stops.Count - 1].Offset, 6);

        HtmlRenderDocument bounded = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 70D,
            Margins = HtmlRenderMargins.All(8D),
            MaxGradientStops = 4
        });
        Assert.DoesNotContain(bounded.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.FillRadialGradient != null);
        Assert.Contains(bounded.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GradientStopLimitExceeded);
    }

    [Fact]
    public void HtmlConicGradient_RemainsVectorAcrossPngSvgAndSearchablePdf() {
        const string html = "<div style='width:80px;height:60px;border-radius:8px;background:conic-gradient(from 0deg at 25% 50%,red 0 25%,blue 25% 75%,lime 75% 100%)'>ConicMarker</div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 120D,
            Margins = HtmlRenderMargins.All(8D),
            ConicGradientQualitySegments = 72
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderPathClipGroup clipped = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderPathClipGroup>());
        HtmlRenderDrawing visual = Assert.Single(clipped.Visuals.OfType<HtmlRenderDrawing>());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = HtmlConversionDocument.Parse(html).ToSvg(options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions(options));

        Assert.NotEmpty(visual.Drawing.Elements);
        Assert.NotEqual(raster.GetPixel(28, 10), raster.GetPixel(86, 38));
        Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
        Assert.True(CountBackgroundOccurrences(svg, "<path") >= 72);
        Assert.Contains("ConicMarker", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
        Assert.DoesNotContain(HtmlConversionDocument.Parse(html).ToPdfDocumentResult(new HtmlPdfSaveOptions(options)).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlGradientCentersResolveThreeAndFourComponentEdgeOffsets() {
        Assert.True(HtmlCssConicGradientParser.TryParse(
            "conic-gradient(at right 10px bottom 20px,red,blue)",
            8,
            out HtmlCssConicGradientDefinition? conicDefinition,
            out bool conicLimitExceeded));
        Assert.NotNull(conicDefinition);
        Assert.False(conicLimitExceeded);
        Assert.True(conicDefinition!.TryResolve(100D, 80D, 16D, 16D, 100D, 80D, out OfficeConicGradient? conic, out _));
        Assert.NotNull(conic);
        Assert.Equal(0.9D, conic!.CenterX, 6);
        Assert.Equal(0.75D, conic.CenterY, 6);

        Assert.True(HtmlCssRadialGradientParser.TryParse(
            "radial-gradient(circle at left 15px bottom,red,blue)",
            8,
            out HtmlCssRadialGradientDefinition? radialDefinition,
            out bool radialLimitExceeded));
        Assert.NotNull(radialDefinition);
        Assert.False(radialLimitExceeded);
        Assert.True(radialDefinition!.TryResolve(100D, 80D, 16D, 16D, 100D, 80D, out OfficeRadialGradient? radial, out _));
        Assert.NotNull(radial);
        Assert.Equal(0.15D, radial!.EndX, 6);
        Assert.Equal(1D, radial.EndY, 6);

        Assert.False(HtmlCssConicGradientParser.TryParse(
            "conic-gradient(at right 10px left 20px,red,blue)",
            8,
            out _,
            out _));
    }

    [Fact]
    public void HtmlConicGradient_StylesheetShorthandFlowsThroughLayeredCascadeAndExports() {
        const string html = """
            <style>
              @layer base, theme;
              @layer base { .card { background:red; } }
              @layer theme { .card { background:conic-gradient(from 0deg at 25% 50%,red 0 25%,blue 25% 75%,lime 75% 100%); } }
            </style>
            <div class="card" style="width:80px;height:60px">StylesheetConic</div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 120D,
            Margins = HtmlRenderMargins.All(8D),
            BackgroundColor = OfficeColor.Transparent,
            ConicGradientQualitySegments = 12
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderDrawing visual = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = HtmlConversionDocument.Parse(html).ToSvg(options);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions(options));

        Assert.NotEmpty(visual.Drawing.Elements);
        Assert.NotEqual(OfficeColor.Transparent, raster.GetPixel(8, 8));
        Assert.NotEqual(OfficeColor.Transparent, raster.GetPixel(87, 67));
        Assert.True(CountBackgroundOccurrences(svg, "<path") >= 12);
        Assert.Contains("StylesheetConic", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlGradients_ResolveContainerUnitsAgainstTheActiveQueryContainer() {
        const string html = "<section style='width:200px;height:100px;container-type:size'>"
            + "<div style='width:100px;height:20px;background:linear-gradient(to right,red 0,blue 10cqw)'></div>"
            + "<div style='width:100px;height:50px;background:radial-gradient(circle 10cqw at 10cqw 10cqh,red,blue)'></div>"
            + "</section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 500D,
            ViewportHeight = 300D,
            Margins = HtmlRenderMargins.All(0D)
        });
        OfficeLinearGradient linear = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        OfficeRadialGradient radial = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillRadialGradient != null).Shape.FillRadialGradient!;

        Assert.Contains(linear.Stops, stop => Math.Abs(stop.Offset - 0.2D) < 0.0001D);
        Assert.Equal(0.2D, radial.EndX, 6);
        Assert.Equal(0.2D, radial.EndY, 6);
        Assert.Equal(0.2D, radial.EndRadiusX, 6);
        Assert.Equal(0.4D, radial.EndRadiusY, 6);

        Assert.True(HtmlCssConicGradientParser.TryParse(
            "conic-gradient(at 10cqw 10cqh,red,blue)",
            8,
            out HtmlCssConicGradientDefinition? definition,
            out bool stopLimitExceeded));
        Assert.NotNull(definition);
        Assert.False(stopLimitExceeded);
        Assert.True(definition!.TryResolve(100D, 50D, 16D, 16D, 500D, 300D, 200D, 100D, out OfficeConicGradient? conic, out _));
        Assert.NotNull(conic);
        Assert.Equal(0.2D, conic!.CenterX, 6);
        Assert.Equal(0.2D, conic.CenterY, 6);
    }

    [Fact]
    public void HtmlConicGradient_AcceptsCssColor4Stops() {
        const string gradient = "conic-gradient(from 210deg at 80% 20%,oklch(72% .18 155),oklch(62% .2 245),oklch(68% .2 320),oklch(72% .18 155))";
        string html = "<style>.hero{width:80px;height:60px;background:" + gradient + "}</style><div class='hero'></div>";

        Assert.True(HtmlCssConicGradientParser.TryParse(gradient, 32, out HtmlCssConicGradientDefinition? definition, out bool stopLimitExceeded));
        Assert.NotNull(definition);
        Assert.False(stopLimitExceeded);
        Assert.True(HtmlCssConicGradientParser.TryParse(
            "conic-gradient(from 210deg, at 80% 20%, rgba(0, 196, 112, 1), rgba(0, 139, 245, 1))",
            32,
            out _,
            out _));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(8D),
            ConicGradientQualitySegments = 72
        });
        HtmlRenderDrawing visual = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 80D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(8D)
        });
        OfficeDrawing pdfDrawing = PdfCore.PdfPageImageRenderer.RenderPage(pdf);

        Assert.NotEmpty(visual.Drawing.Elements);
        Assert.NotEqual(raster.GetPixel(20, 20), raster.GetPixel(70, 50));
        Assert.Contains(pdfDrawing.Shapes, shape => shape.Shape.FillColor.HasValue && shape.Shape.FillColor.Value.G > shape.Shape.FillColor.Value.R);
        Assert.Contains(pdfDrawing.Shapes, shape => shape.Shape.FillColor.HasValue && shape.Shape.FillColor.Value.B > shape.Shape.FillColor.Value.G);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlRepeatingConicGradient_ExpandsAngularStopsAndHonorsTheStopBudget() {
        const string background = "repeating-conic-gradient(from .25turn at center,red 0 30deg,blue 30deg 60deg)";
        string html = "<div style='width:40px;height:40px;background:" + background + "'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 70D,
            Margins = HtmlRenderMargins.All(8D),
            MaxGradientStops = 64,
            ConicGradientQualitySegments = 36
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderDrawing drawing = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        Assert.NotEmpty(drawing.Drawing.Elements);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);

        options.MaxGradientStops = 4;
        HtmlRenderDocument bounded = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        Assert.DoesNotContain(bounded.Pages[0].Visuals, visual => visual is HtmlRenderDrawing);
        Assert.Contains(bounded.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GradientStopLimitExceeded);
    }

    [Fact]
    public void HtmlLinearGradient_ClipsOutOfRangeStopsInsteadOfDroppingThePaint() {
        const string html = "<div style='width:100px;height:20px;background:linear-gradient(to right,red -50%,lime 50%,blue 150%)'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 130D,
            Margins = HtmlRenderMargins.All(8D)
        });

        OfficeLinearGradient gradient = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null).Shape.FillGradient!;
        Assert.Equal(new[] { 0D, 0.5D, 1D }, gradient.Stops.Select(stop => stop.Offset));
        Assert.Equal(OfficeColor.FromRgb(128, 128, 0), gradient.Stops[0].Color);
        Assert.Equal(OfficeColor.FromRgb(0, 128, 128), gradient.Stops[2].Color);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void HtmlLinearGradient_ComposesAboveUrlLayers() {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(0, 0, 255));
        string html = "<div style=\"width:40px;height:20px;background-image:linear-gradient(to right,red,lime),url('data:image/png;base64,"
            + imageData
            + "');background-size:auto,40px 20px;background-repeat:no-repeat\"></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
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

    [Fact]
    public void HtmlLinearGradient_OversizedBoxKeepsOneContinuousPaintAcrossPageClips() {
        const string html = "<html style='margin:0'><body style='margin:0'><div id='tall-gradient' style='width:40px;height:90px;background:linear-gradient(to bottom,red,blue)'></div></body></html>";
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
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => {
            HtmlRenderClipGroup fragment = Assert.Single(page.Visuals.OfType<HtmlRenderClipGroup>(), group =>
                group.Visuals.OfType<HtmlRenderShape>().Any(shape => shape.Shape.FillGradient != null));
            Assert.Single(fragment.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.FillGradient != null);
        });
        Assert.True(OfficePngReader.TryDecode(pngPages[0].Bytes, out OfficeRasterImage? first));
        Assert.True(OfficePngReader.TryDecode(pngPages[1].Bytes, out OfficeRasterImage? second));
        Assert.True(OfficePngReader.TryDecode(pngPages[2].Bytes, out OfficeRasterImage? third));
        Assert.True(first!.GetPixel(20, 20).R > first.GetPixel(20, 20).B);
        Assert.True(second!.GetPixel(20, 20).B > second.GetPixel(20, 20).R);
        Assert.True(third!.GetPixel(20, 5).B > third.GetPixel(20, 5).R);
        Assert.All(svgPages, page => {
            string svg = Encoding.UTF8.GetString(page.Bytes);
            Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
            Assert.Contains("<linearGradient", svg, StringComparison.Ordinal);
        });
        Assert.Equal(3, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Theory]
    [InlineData("linear-gradient(red,lime,blue)")]
    [InlineData("radial-gradient(red,lime,blue)")]
    public void HtmlRender_BoundsCssGradientStops(string background) {
        string html = "<div style='width:40px;height:20px;background:" + background + "'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            Margins = HtmlRenderMargins.All(8D),
            MaxGradientStops = 2
        });

        Assert.DoesNotContain(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillGradient != null || shape.Shape.FillRadialGradient != null);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GradientStopLimitExceeded);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Fact]
    public void CssLengthAndGradientParsingRejectNonFiniteResults() {
        Assert.False(HtmlRenderCssValues.TryLength(
            "1e308cm",
            reference: 100D,
            fontSize: 16D,
            rootFontSize: 16D,
            out _));
        Assert.False(HtmlRenderCssValues.TryLength(
            "1e308em",
            reference: 100D,
            fontSize: 16D,
            rootFontSize: 16D,
            out _));
        Assert.False(HtmlCssGradientStops.TryParse(
            new[] { "red 0px", "blue 1e308cm" },
            startIndex: 0,
            maximumStops: 8,
            out _,
            out _));
    }
}
