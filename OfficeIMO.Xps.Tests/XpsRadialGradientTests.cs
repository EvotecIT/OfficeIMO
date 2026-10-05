using System;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsRadialGradientTests {
    [Theory]
    [InlineData(XpsFormat.Xps, "0.8,0.6,-0.6,0.8,80,-20", false)]
    [InlineData(XpsFormat.OpenXps, "1,0.3,0.4,1,-30,-20", false)]
    [InlineData(XpsFormat.Xps, "1,0.3,0.4,1,-30,-20", true)]
    [InlineData(XpsFormat.OpenXps, "-1,0.3,0.4,1,150,-20", true)]
    public void AffineRadialBrushMatchesItsAnalyticFieldAndSurvivesPdf(XpsFormat format, string matrix, bool alpha) {
        var doc = Create(format, matrix, alpha); var page = doc.Pages[0];
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var image));
        var numbers = matrix.Split(',').Select(x => double.Parse(x, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
        var inverse = new OfficeTransform(numbers[0], numbers[1], numbers[2], numbers[3], numbers[4], numbers[5]).Invert();
        foreach (var xy in new[] { (70, 60), (100, 80), (120, 90), (140, 60) }) {
            var point = inverse.TransformPoint(new OfficePoint(xy.Item1 + .5, xy.Item2 + .5));
            double ratio = Math.Min(1, Math.Sqrt(Math.Pow((point.X - 100) / 60, 2) + Math.Pow((point.Y - 80) / 25, 2)));
            double opacity = alpha ? (64 + (191 * ratio)) / 255 : 1;
            var pixel = image!.GetPixel(xy.Item1, xy.Item2);
            Assert.InRange(Math.Abs(pixel.R - (255 * (1 - ratio) * opacity + 255 * (1 - opacity))), 0, 4);
            Assert.InRange(Math.Abs(pixel.B - (255 * ratio * opacity + 255 * (1 - opacity))), 0, 4);
        }
        var pdfPage = Assert.Single(PdfReadDocument.Open(doc.ToPdf()).Pages);
        var drawing = page.ToDrawing();
        var svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.True(OfficeSvgDrawingReader.TryRead(System.Text.Encoding.UTF8.GetBytes(svg), out var roundtrip));
        Assert.NotNull(roundtrip);
        var svgImage = OfficeDrawingRasterRenderer.Render(roundtrip!, scale: .75, background: OfficeColor.White);
        foreach (var xy in new[] { (70, 60), (100, 80), (120, 90) }) {
            Assert.InRange(Math.Abs(image!.GetPixel(xy.Item1, xy.Item2).R - svgImage.GetPixel(xy.Item1, xy.Item2).R), 0, 2);
        }
        if (!alpha) {
            var pdfImage = OfficeDrawingRasterRenderer.Render(pdfPage.ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
            foreach (var xy in new[] { (70, 60), (100, 80), (120, 90) }) {
                Assert.InRange(Math.Abs(image!.GetPixel(xy.Item1, xy.Item2).R - pdfImage.GetPixel(xy.Item1, xy.Item2).R), 0, 3);
            }
        }
    }

    [Theory]
    [InlineData("Repeat")]
    [InlineData("Reflect")]
    public void AffineRadialSpreadRepeatsTheNativeColorField(string spread) {
        var doc = Create(XpsFormat.OpenXps, "1,0.3,0.4,1,-30,-20", false, spread);
        Assert.True(OfficeRasterImageDecoder.TryDecode(doc.Pages[0].ExportImage(OfficeImageExportFormat.Png).Bytes, out var image));
        var inverse = new OfficeTransform(1,.3,.4,1,-30,-20).Invert();
        foreach (var xy in new[] { (10, 10), (30, 45), (160, 120), (190, 150) }) {
            var point = inverse.TransformPoint(new OfficePoint(xy.Item1 + .5, xy.Item2 + .5));
            double ratio = Math.Sqrt(Math.Pow((point.X - 100) / 60, 2) + Math.Pow((point.Y - 80) / 25, 2));
            ratio %= spread == "Repeat" ? 1 : 2;
            if (ratio > 1) ratio = 2 - ratio;
            Assert.InRange(Math.Abs(image!.GetPixel(xy.Item1, xy.Item2).R - 255 * (1 - ratio)), 0, 3);
        }
        Assert.Single(PdfReadDocument.Open(doc.ToPdf()).Pages);
    }

    [Theory]
    [InlineData("Pad")]
    [InlineData("Repeat")]
    [InlineData("Reflect")]
    public void AffineRadialStrokePreservesTheBrushCoordinateSystem(string spread) {
        var doc = Create(XpsFormat.Xps, "1,0.3,0.4,1,-30,-20", false, spread); var page = doc.Pages[0];
        var xml = page.GetMarkup(); var ns = xml.Name.Namespace; var path = xml.Element(ns + "Path")!;
        path.SetAttributeValue("Data", "M30,50L170,100"); path.SetAttributeValue("StrokeThickness", "20");
        path.Element(ns + "Path.Fill")!.Name = ns + "Path.Stroke"; page.ReplaceMarkup(xml);
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var image));
        var inverse = new OfficeTransform(1,.3,.4,1,-30,-20).Invert();
        foreach (var xy in new[] { (60, 61), (100, 75), (140, 89) }) {
            var point = inverse.TransformPoint(new OfficePoint(xy.Item1 + .5, xy.Item2 + .5));
            double ratio = Math.Sqrt(Math.Pow((point.X - 100) / 60, 2) + Math.Pow((point.Y - 80) / 25, 2));
            ratio = spread == "Pad" ? Math.Min(1, ratio) : ratio % (spread == "Repeat" ? 1 : 2);
            if (ratio > 1) ratio = 2 - ratio;
            Assert.InRange(Math.Abs(image!.GetPixel(xy.Item1, xy.Item2).R - 255 * (1 - ratio)), 0, 3);
        }
        Assert.Single(PdfReadDocument.Open(doc.ToPdf()).Pages);
    }

    [Fact]
    public void HighFrequencyRadialSpreadSamplesWithoutStopExpansion() {
        var doc = Create(XpsFormat.Xps, "1,0,0,1,0,0", false, "Repeat"); var page = doc.Pages[0];
        var xml = page.GetMarkup(); var brush = xml.Descendants().Single(e => e.Name.LocalName == "RadialGradientBrush");
        brush.SetAttributeValue("RadiusX", ".001"); brush.SetAttributeValue("RadiusY", ".001"); page.ReplaceMarkup(xml);
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var image));
        double ratio = Math.Sqrt(.5 * .5 + .5 * .5) / .001;
        ratio -= Math.Floor(ratio);
        Assert.InRange(Math.Abs(image!.GetPixel(100, 80).B - 255 * ratio), 0, 2);
        XpsUnboundedRadialSpreadTests.AssertPdfMatchesNative(doc, (100, 80), (50, 20), (180, 140));
    }

    internal static XpsDocument Create(XpsFormat format, string matrix, bool alpha, string spread = "Pad") {
        var doc = XpsDocument.Create(format); var page = doc.AddPage(200, 160);
        var xml = page.GetMarkup(); var ns = xml.Name.Namespace;
        xml.Add(new XElement(ns + "Path", new XAttribute("Data", "M0,0H200V160H0Z"),
            new XElement(ns + "Path.Fill", new XElement(ns + "RadialGradientBrush", new XAttribute("Center", "100,80"), new XAttribute("GradientOrigin", "100,80"),
                new XAttribute("SpreadMethod", spread), new XAttribute("RadiusX", "60"), new XAttribute("RadiusY", "25"), new XAttribute("Transform", matrix),
                new XElement(ns + "RadialGradientBrush.GradientStops",
                    new XElement(ns + "GradientStop", new XAttribute("Offset", "0"), new XAttribute("Color", alpha ? "#40FF0000" : "#FFFF0000")),
                    new XElement(ns + "GradientStop", new XAttribute("Offset", "1"), new XAttribute("Color", "#FF0000FF")))))));
        page.ReplaceMarkup(xml); return doc;
    }
}
