using System;
using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsRadialExteriorSpreadTests {
    [Theory]
    [InlineData(XpsFormat.Xps, "Repeat", "1,0,0,1,0,0", false, false)]
    [InlineData(XpsFormat.OpenXps, "Repeat", "1,0.3,0.4,1,-30,-20", true, false)]
    [InlineData(XpsFormat.Xps, "Repeat", "1,0.3,0.4,1,-30,-20", false, true)]
    [InlineData(XpsFormat.OpenXps, "Repeat", "-1,0.3,0.4,1,150,-20", false, false)]
    [InlineData(XpsFormat.Xps, "Reflect", "1,0,0,1,0,0", false, false)]
    [InlineData(XpsFormat.OpenXps, "Reflect", "1,0.3,0.4,1,-30,-20", true, false)]
    [InlineData(XpsFormat.Xps, "Reflect", "1,0.3,0.4,1,-30,-20", false, true)]
    [InlineData(XpsFormat.OpenXps, "Reflect", "-1,0.3,0.4,1,150,-20", false, false)]
    public void NativeExteriorSpreadUsesFirstEllipseAndCorrectOutsideColor(XpsFormat format, string spread, string matrix, bool alpha, bool stroke) {
        var doc = Create(format, spread, matrix, alpha, stroke); var page = XpsDocument.Load(doc.Save()).Pages[0];
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var image));
        var values = matrix.Split(',').Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        var inverse = new OfficeTransform(values[0], values[1], values[2], values[3], values[4], values[5]).Invert();
        var points = stroke ? new[] { (60, 61), (100, 75), (140, 89) } : new[] { (10, 10), (30, 45), (100, 80), (120, 80), (170, 120) };
        foreach (var xy in points) {
            var point = inverse.TransformPoint(new OfficePoint(xy.Item1 + .5, xy.Item2 + .5));
            double ratio = NativeRatio(point.X, point.Y, spread);
            double opacity = alpha ? (64 + 64 * ratio) / 255D : 1D;
            var pixel = image!.GetPixel(xy.Item1, xy.Item2);
            Assert.InRange(Math.Abs(pixel.R - (255 * (1 - ratio) * opacity + 255 * (1 - opacity))), 0, 4);
            Assert.InRange(Math.Abs(pixel.G - 255 * (1 - opacity)), 0, 4);
            Assert.InRange(Math.Abs(pixel.B - (255 * ratio * opacity + 255 * (1 - opacity))), 0, 4);
        }
        var pdfPage = Assert.Single(PdfReadDocument.Open(doc.ToPdf()).Pages);
        if (!alpha && !stroke) {
            var readback = OfficeDrawingRasterRenderer.Render(pdfPage.ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
            foreach (var xy in points) {
                var expected = image!.GetPixel(xy.Item1, xy.Item2); var actual = readback.GetPixel(xy.Item1, xy.Item2);
                Assert.InRange(Math.Abs(actual.R - expected.R), 0, 4);
                Assert.InRange(Math.Abs(actual.G - expected.G), 0, 4);
                Assert.InRange(Math.Abs(actual.B - expected.B), 0, 4);
            }
        }
        AssertDirectSvgMatches(page, image!, points);
        Assert.Contains("<pattern", OfficeDrawingSvgExporter.ToSvg(page.ToDrawing()));
    }

    [Fact]
    public void NearBoundaryExteriorBeyondExpansionBudgetUsesExplicitSpread() {
        var doc = Create(XpsFormat.OpenXps, "Repeat", "1,0,0,1,0,0", false, false);
        var page = doc.Pages[0]; var xml = page.GetMarkup();
        xml.Descendants().Single(e => e.Name.LocalName == "RadialGradientBrush").SetAttributeValue("GradientOrigin", "160.00000001,80");
        page.ReplaceMarkup(xml);
        Assert.NotNull(page.ToDrawing());
        XpsUnboundedRadialSpreadTests.AssertPdfMatchesNative(doc, (20, 20), (140, 30), (159, 120), (180, 80));
        Assert.Empty(page.ToSvg().Diagnostics);
    }

    [Theory]
    [InlineData(XpsFormat.Xps, "Repeat")]
    [InlineData(XpsFormat.OpenXps, "Reflect")]
    public void DirectSpreadCoversUnfilledStrokedFigures(XpsFormat format, string spread) {
        var document = XpsFigurePaintTests.Create(format, false);
        var page = document.Pages[0]; var xml = page.GetMarkup(); var ns = xml.Name.Namespace;
        var path = xml.Element(ns + "Path")!;
        path.Attribute("Stroke")!.Remove();
        path.SetAttributeValue("StrokeThickness", "4"); path.SetAttributeValue("StrokeMiterLimit", "1");
        var brush = Create(format, spread, "1,0,0,1,0,0", false, false).Pages[0].GetMarkup()
            .Descendants().Single(element => element.Name.LocalName == "RadialGradientBrush");
        path.Add(new System.Xml.Linq.XElement(ns + "Path.Stroke", brush)); page.ReplaceMarkup(xml);
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var image));
        AssertDirectSvgMatches(page, image!, new[] { (20, 40), (80, 50), (110, 50), (170, 50) });
    }

    internal static void AssertDirectSvgMatches(XpsPage page, OfficeRasterImage expected, (int, int)[] points) {
        var svg = page.ToSvg();
        Assert.Empty(svg.Diagnostics);
        Assert.True(OfficeSvgDrawingReader.TryRead(System.Text.Encoding.UTF8.GetBytes(svg.Svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        var actual = OfficeDrawingRasterRenderer.Render(drawing!, background: OfficeColor.White);
        foreach (var xy in points) {
            var left = expected.GetPixel(xy.Item1, xy.Item2); var right = actual.GetPixel(xy.Item1, xy.Item2);
            Assert.InRange(Math.Abs(left.R - right.R), 0, 4);
            Assert.InRange(Math.Abs(left.G - right.G), 0, 4);
            Assert.InRange(Math.Abs(left.B - right.B), 0, 4);
        }
    }

    internal static XpsDocument Create(XpsFormat format, string spread, string matrix, bool alpha, bool stroke) {
        var doc = XpsRadialBoundaryTests.Create(format, 200, matrix, alpha, stroke);
        var page = doc.Pages[0]; var xml = page.GetMarkup();
        xml.Descendants().Single(e => e.Name.LocalName == "RadialGradientBrush").SetAttributeValue("SpreadMethod", spread);
        page.ReplaceMarkup(xml); return doc;
    }

    private static double NativeRatio(double x, double y, string spread) {
        double origin = 100D / 60D, px = (x - 200) / 60D, py = (y - 80) / 25D;
        double a = origin * origin - 1, b = 2 * px * origin, c = px * px + py * py, d = b * b - 4 * a * c;
        if (d < 0 || b >= 0) return spread == "Reflect" ? 0 : 1;
        double ratio = 2 * c / (-b + Math.Sqrt(d));
        ratio %= spread == "Reflect" ? 2 : 1;
        return ratio > 1 ? 2 - ratio : ratio;
    }
}
