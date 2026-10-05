using System;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsUnboundedRadialSpreadTests {
    [Theory]
    [InlineData(XpsFormat.Xps, "Repeat", false, false)]
    [InlineData(XpsFormat.OpenXps, "Repeat", true, true)]
    [InlineData(XpsFormat.Xps, "Reflect", true, false)]
    [InlineData(XpsFormat.OpenXps, "Reflect", false, true)]
    public void UnboundedFieldSurvivesNativeRasterAndBothSvgRoutes(XpsFormat format, string spread, bool alpha, bool linear) {
        var doc = XpsRadialBoundaryTests.Create(format, 160, "1,0,0,1,0,0", alpha, false);
        var page = doc.Pages[0]; var xml = page.GetMarkup();
        var brush = xml.Descendants().Single(e => e.Name.LocalName == "RadialGradientBrush");
        brush.SetAttributeValue("SpreadMethod", spread);
        brush.SetAttributeValue("ColorInterpolationMode", linear ? "ScRgbLinearInterpolation" : "SRgbLinearInterpolation");
        page.ReplaceMarkup(xml);
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var native));
        string direct = page.ToSvg().Svg;
        string drawingSvg = OfficeDrawingSvgExporter.ToSvg(page.ToDrawing(), 1, OfficeSvgSizeUnit.Pixel);
        var images = new System.Collections.Generic.List<OfficeRasterImage> { native! };
        foreach (string svg in new[] { direct, drawingSvg }) {
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var imported, out int unsupported));
            Assert.Equal(0, unsupported);
            images.Add(OfficeDrawingRasterRenderer.Render(imported!, background: OfficeColor.White));
        }
        foreach (var xy in new[] { (10, 10), (140, 30), (159, 120), (180, 80) }) {
            double px = (xy.Item1 + .5 - 160) / 60D, py = (xy.Item2 + .5 - 80) / 25D;
            double q = spread == "Repeat" ? 1 : 0;
            if (px < 0) { q = -(px * px + py * py) / (2 * px); q %= spread == "Repeat" ? 1 : 2; if (q > 1) q = 2 - q; }
            double opacity = alpha ? (64 + 64 * q) / 255D : 1;
            double Encode(double c) => linear ? (c <= .0031308 ? 12.92 * c : 1.055 * Math.Pow(c, 1 / 2.4) - .055) : c;
            foreach (var image in images) {
                var pixel = image.GetPixel(xy.Item1, xy.Item2);
                Assert.InRange(Math.Abs(pixel.R - 255 * (Encode(1 - q) * opacity + 1 - opacity)), 0, 4);
                Assert.InRange(Math.Abs(pixel.G - 255 * (1 - opacity)), 0, 4);
                Assert.InRange(Math.Abs(pixel.B - 255 * (Encode(q) * opacity + 1 - opacity)), 0, 4);
            }
        }
    }
}
