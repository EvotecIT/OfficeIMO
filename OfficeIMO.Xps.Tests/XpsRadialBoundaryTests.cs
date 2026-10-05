using System;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsRadialBoundaryTests {
    [Theory]
    [InlineData(XpsFormat.Xps, 160, "1,0,0,1,0,0", false, false)]
    [InlineData(XpsFormat.OpenXps, 160, "1,0,0,1,0,0", false, false)]
    [InlineData(XpsFormat.Xps, 200, "1,0,0,1,0,0", false, false)]
    [InlineData(XpsFormat.OpenXps, 200, "1,0,0,1,0,0", false, false)]
    [InlineData(XpsFormat.Xps, 200, "1,0.3,0.4,1,-30,-20", false, true)]
    [InlineData(XpsFormat.OpenXps, 200, "1,0.3,0.4,1,-30,-20", true, false)]
    [InlineData(XpsFormat.Xps, 160, "0.8,0.6,-0.6,0.8,80,-20", true, false)]
    [InlineData(XpsFormat.OpenXps, 160, "-1,0.3,0.4,1,150,-20", false, true)]
    public void NativePadUsesSmallestContainingEllipseAndOutsideEndpoint(XpsFormat format, int focus, string matrix, bool alpha, bool stroke) {
        var document = Create(format, focus, matrix, alpha, stroke);
        var page = XpsDocument.Load(document.Save()).Pages[0];
        var values = matrix.Split(',').Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        var inverse = new OfficeTransform(values[0], values[1], values[2], values[3], values[4], values[5]).Invert();
        var raster = Raster(page);
        var points = stroke ? new[] { (60,61), (100,75), (140,89) } : new[] { (100,80), (120,80), (140,80), (170,80), (40,10) };
        foreach (var xy in points) {
            var point = inverse.TransformPoint(new OfficePoint(xy.Item1 + .5, xy.Item2 + .5));
            double ratio = SmallestEllipse(point.X, point.Y, focus);
            double opacity = alpha ? (64 + 64 * ratio) / 255D : 1D;
            var pixel = raster.GetPixel(xy.Item1, xy.Item2);
            Assert.InRange(Math.Abs(pixel.R - (255 * (1 - ratio) * opacity + 255 * (1 - opacity))), 0, 4);
            Assert.InRange(Math.Abs(pixel.G - 255 * (1 - opacity)), 0, 4);
            Assert.InRange(Math.Abs(pixel.B - (255 * ratio * opacity + 255 * (1 - opacity))), 0, 4);
        }
        byte[] pdf = document.ToPdf();
        Assert.Single(PdfReadDocument.Open(pdf).Pages);
        var svg = page.ToSvg();
        Assert.Empty(svg.Diagnostics);
        var svgXml = XElement.Parse(svg.Svg);
        var fields = svgXml.Descendants().Where(e => e.Name.LocalName == "radialGradient").ToArray();
        Assert.Equal(2, fields.Length);
        Assert.All(fields, field => {
            Assert.Equal("0", (string?)field.Attribute("r"));
            Assert.Equal("60", (string?)field.Attribute("fr"));
            Assert.Equal("100", (string?)field.Attribute("fx"));
            Assert.Equal(focus.ToString(CultureInfo.InvariantCulture), (string?)field.Attribute("cx"));
            Assert.All(field.Elements(), stop => Assert.Null(stop.Attribute("stop-opacity")));
        });
        Assert.Contains("<pattern", svg.Svg);
        Assert.True(OfficeSvgDrawingReader.TryRead(System.Text.Encoding.UTF8.GetBytes(svg.Svg), out var imported, out int unsupported));
        Assert.Equal(0, unsupported);
        var importedRaster = OfficeDrawingRasterRenderer.Render(imported!, background: OfficeColor.White);
        foreach (var xy in points) {
            var expected = raster.GetPixel(xy.Item1, xy.Item2); var actual = importedRaster.GetPixel(xy.Item1, xy.Item2);
            Assert.InRange(Math.Abs(actual.R - expected.R), 0, 4);
            Assert.InRange(Math.Abs(actual.G - expected.G), 0, 4);
            Assert.InRange(Math.Abs(actual.B - expected.B), 0, 4);
        }
        Assert.NotEmpty(page.ExportImage(OfficeImageExportFormat.Svg).Bytes);
        Assert.Contains("<pattern", OfficeDrawingSvgExporter.ToSvg(page.ToDrawing()));
        if (!alpha && !stroke) {
            var readback = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
            foreach (var xy in points.Concat(new[] { (100, 100) })) {
                var expected = raster.GetPixel(xy.Item1, xy.Item2); var actual = readback.GetPixel(xy.Item1, xy.Item2);
                Assert.InRange(Math.Abs(actual.R - expected.R), 0, 4);
                Assert.InRange(Math.Abs(actual.G - expected.G), 0, 4);
                Assert.InRange(Math.Abs(actual.B - expected.B), 0, 4);
            }
        }
    }

    [Theory]
    [InlineData("Repeat")]
    [InlineData("Reflect")]
    public void UnboundedBoundarySpreadExportsSvgWhileDrawingRemainsExplicit(string spread) {
        var document = Create(XpsFormat.OpenXps, 160, "1,0,0,1,0,0", false, false);
        var page = document.Pages[0]; var markup = page.GetMarkup();
        markup.Descendants().Single(element => element.Name.LocalName == "RadialGradientBrush").SetAttributeValue("SpreadMethod", spread);
        page.ReplaceMarkup(markup);
        Assert.Throws<NotSupportedException>(() => page.ToDrawing());
        Assert.Throws<NotSupportedException>(() => document.ToPdf());
        var result = page.ToSvg();
        Assert.Empty(result.Diagnostics);
        var fields = XElement.Parse(result.Svg).Descendants().Where(e => e.Name.LocalName == "radialGradient").ToArray();
        Assert.Equal(2, fields.Length);
        Assert.All(fields, field => {
            Assert.Equal(spread.ToLowerInvariant(), (string?)field.Attribute("spreadMethod"));
            Assert.Equal("0", (string?)field.Attribute("r"));
        });
    }

    [Theory]
    [InlineData(XpsFormat.Xps, "Repeat", false, 140)]
    [InlineData(XpsFormat.Xps, "Repeat", false, 128)]
    [InlineData(XpsFormat.OpenXps, "Repeat", true, 140)]
    [InlineData(XpsFormat.Xps, "Reflect", true, 140)]
    [InlineData(XpsFormat.OpenXps, "Reflect", false, 140)]
    [InlineData(XpsFormat.OpenXps, "Reflect", false, 128)]
    public void BoundedBoundarySpreadRetainsNativeField(XpsFormat format, string spread, bool alpha, int right) {
        var document = Create(format, 160, "1,0,0,1,0,0", alpha, false);
        var page = document.Pages[0]; var markup = page.GetMarkup();
        markup.Descendants().Single(e => e.Name.LocalName == "Path").SetAttributeValue("Data", right == 128 ? "M0,10H128V150H0Z" : "M10,10H140V150H10Z");
        markup.Descendants().Single(e => e.Name.LocalName == "RadialGradientBrush").SetAttributeValue("SpreadMethod", spread);
        page.ReplaceMarkup(markup);
        var raster = Raster(XpsDocument.Load(document.Save()).Pages[0]);
        foreach (var xy in new[] { (20, 20), (60, 40), (100, 80), (right - 10, 130) }) {
            double px = (xy.Item1 + .5 - 160) / 60D, py = (xy.Item2 + .5 - 80) / 25D;
            double ratio = -(px * px + py * py) / (2 * px);
            ratio %= spread == "Repeat" ? 1 : 2;
            if (ratio > 1) ratio = 2 - ratio;
            double opacity = alpha ? (64 + 64 * ratio) / 255D : 1;
            var pixel = raster.GetPixel(xy.Item1, xy.Item2);
            Assert.InRange(Math.Abs(pixel.R - (255 * (1 - ratio) * opacity + 255 * (1 - opacity))), 0, 4);
            Assert.InRange(Math.Abs(pixel.G - 255 * (1 - opacity)), 0, 4);
            Assert.InRange(Math.Abs(pixel.B - (255 * ratio * opacity + 255 * (1 - opacity))), 0, 4);
        }
        XpsRadialExteriorSpreadTests.AssertDirectSvgMatches(page, raster,
            new[] { (20, 20), (60, 40), (100, 80), (right - 10, 130) });
        var pdf = PdfReadDocument.Open(document.ToPdf());
        Assert.Single(pdf.Pages);
        if (!alpha) {
            var readback = OfficeDrawingRasterRenderer.Render(pdf.Pages[0].ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
            foreach (var xy in new[] { (20, 20), (60, 40), (100, 80), (right - 10, 130) }) {
                Assert.InRange(Math.Abs(readback.GetPixel(xy.Item1, xy.Item2).R - raster.GetPixel(xy.Item1, xy.Item2).R), 0, 4);
            }
        }
    }

    [Theory]
    [InlineData("Repeat")]
    [InlineData("Reflect")]
    public void BoundedBoundarySpreadStillEnforcesStopBudget(string spread) {
        var doc = Create(XpsFormat.OpenXps, 160, "1,0,0,1,0,0", false, false);
        var page = doc.Pages[0]; var xml = page.GetMarkup();
        xml.Descendants().Single(e => e.Name.LocalName == "Path").SetAttributeValue("Data", "M10,10H159.999V150H10Z");
        xml.Descendants().Single(e => e.Name.LocalName == "RadialGradientBrush").SetAttributeValue("SpreadMethod", spread);
        page.ReplaceMarkup(xml);
        Assert.Throws<NotSupportedException>(() => page.ToDrawing());
        Assert.Throws<NotSupportedException>(() => doc.ToPdf());
        var result = page.ToSvg();
        Assert.Empty(result.Diagnostics);
        var fields = XElement.Parse(result.Svg).Descendants().Where(e => e.Name.LocalName == "radialGradient").ToArray();
        Assert.Equal(2, fields.Length);
        Assert.All(fields, field => {
            Assert.Equal(spread.ToLowerInvariant(), (string?)field.Attribute("spreadMethod"));
            Assert.Equal("0", (string?)field.Attribute("r"));
        });
    }

    internal static XpsDocument Create(XpsFormat format, int focus, string matrix, bool alpha, bool stroke) {
        var document = XpsRadialGradientTests.Create(format, matrix, alpha);
        var page = document.Pages[0]; var markup = page.GetMarkup(); var ns = markup.Name.Namespace;
        var brush = markup.Descendants().Single(element => element.Name.LocalName == "RadialGradientBrush");
        brush.SetAttributeValue("GradientOrigin", focus + ",80");
        if (alpha) brush.Descendants().Last(element => element.Name.LocalName == "GradientStop").SetAttributeValue("Color", "#800000FF");
        if (stroke) {
            var path = markup.Element(ns + "Path")!;
            path.SetAttributeValue("Data", "M30,50L170,100"); path.SetAttributeValue("StrokeThickness", "20");
            path.Element(ns + "Path.Fill")!.Name = ns + "Path.Stroke";
        }
        page.ReplaceMarkup(markup); return document;
    }

    private static OfficeRasterImage Raster(XpsPage page) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var raster));
        return raster!;
    }
    private static double SmallestEllipse(double x, double y, double focus) {
        double origin = (focus - 100) / 60D, dx = (x - focus) / 60D, dy = (y - 80) / 25D;
        double a = origin * origin - 1, b = 2 * dx * origin, c = dx * dx + dy * dy;
        if (Math.Abs(a) < 1e-12) return b < 0 ? Math.Min(1, -c / b) : 1;
        double discriminant = b * b - 4 * a * c;
        if (discriminant < 0) return 1;
        double first = (-b - Math.Sqrt(discriminant)) / (2 * a);
        return first < 0 ? 1 : Math.Min(1, first);
    }
}
