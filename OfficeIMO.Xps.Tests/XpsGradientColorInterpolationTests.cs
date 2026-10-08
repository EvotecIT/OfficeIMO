using System;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsGradientColorInterpolationTests {
    [Theory]
    [InlineData(XpsFormat.Xps, "linear", false, "Pad")]
    [InlineData(XpsFormat.OpenXps, "linear", true, "Pad")]
    [InlineData(XpsFormat.Xps, "interior", true, "Pad")]
    [InlineData(XpsFormat.OpenXps, "interior", false, "Pad")]
    [InlineData(XpsFormat.Xps, "boundary", false, "Pad")]
    [InlineData(XpsFormat.OpenXps, "boundary", true, "Pad")]
    [InlineData(XpsFormat.Xps, "exterior", true, "Pad")]
    [InlineData(XpsFormat.OpenXps, "exterior", false, "Pad")]
    [InlineData(XpsFormat.Xps, "exterior", false, "Reflect")]
    [InlineData(XpsFormat.OpenXps, "exterior", true, "Repeat")]
    public void LinearLightColorsSurviveRasterSvgAndPdf(XpsFormat format, string kind, bool alpha, string spread) {
        var document = XpsRadialBoundaryTests.Create(format, kind == "boundary" ? 160 : kind == "exterior" ? 200 : 100, "1,0,0,1,0,0", alpha, false);
        var page = document.Pages[0]; var markup = page.GetMarkup(); var ns = markup.Name.Namespace;
        var brush = markup.Descendants().Single(e => e.Name.LocalName == "RadialGradientBrush");
        brush.SetAttributeValue("ColorInterpolationMode", "ScRgbLinearInterpolation");
        brush.SetAttributeValue("SpreadMethod", spread);
        if (kind == "linear") {
            var stops = brush.Element(ns + "RadialGradientBrush.GradientStops")!;
            stops.Name = ns + "LinearGradientBrush.GradientStops";
            brush.ReplaceWith(new XElement(ns + "LinearGradientBrush", new XAttribute("MappingMode", "Absolute"), new XAttribute("StartPoint", "0,0"), new XAttribute("EndPoint", "200,0"),
                new XAttribute("ColorInterpolationMode", "ScRgbLinearInterpolation"), stops));
        }
        page.ReplaceMarkup(markup);
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var raster));
        var svg = page.ToSvg();
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var imported, out int unsupported));
        Assert.Equal(0, unsupported);
        var direct = OfficeDrawingRasterRenderer.Render(imported!, background: OfficeColor.White);
        var exported = OfficeDrawingSvgExporter.ToSvg(page.ToDrawing(), 1, OfficeSvgSizeUnit.Pixel);
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(exported), out var reimported, out unsupported));
        Assert.Equal(0, unsupported);
        var roundtrip = OfficeDrawingRasterRenderer.Render(reimported!, background: OfficeColor.White);
        var pdf = PdfReadDocument.Open(document.ToPdf());
        var pdfRaster = OfficeDrawingRasterRenderer.Render(pdf.Pages[0].ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
        foreach (int x in new[] { 40, 120, 140 }) {
            double px = x + .5, py = .5 / 25D;
            double ratio;
            if (kind == "linear") ratio = px / 200;
            else if (kind == "interior") ratio = Math.Sqrt(Math.Pow((px - 100) / 60, 2) + py * py);
            else {
                double focus = kind == "boundary" ? 160 : 200;
                double vx = (px - focus) / 60, dx = (focus - 100) / 60;
                double a = dx * dx - 1, b = 2 * vx * dx, c = vx * vx + py * py;
                ratio = a == 0 ? -c / b : 2 * c / (-b + Math.Sqrt(b * b - 4 * a * c));
            }
            if (spread != "Pad") {
                ratio %= spread == "Reflect" ? 2 : 1;
                if (ratio > 1) ratio = 2 - ratio;
            }
            ratio = Math.Max(0, Math.Min(1, ratio));
            double opacity = alpha ? (64 + 64 * ratio) / 255 : 1;
            double Encode(double c) => 255 * (c <= .0031308 ? 12.92 * c : 1.055 * Math.Pow(c, 1 / 2.4) - .055);
            foreach (var image in new[] { raster!, direct, roundtrip, pdfRaster }) {
                var color = image.GetPixel(x, 80);
                Assert.InRange(Math.Abs(color.R - (Encode(1 - ratio) * opacity + 255 * (1 - opacity))), 0, 4);
                Assert.InRange(Math.Abs(color.B - (Encode(ratio) * opacity + 255 * (1 - opacity))), 0, 4);
                Assert.InRange(Math.Abs(color.G - 255 * (1 - opacity)), 0, 4);
            }
        }
    }
}
