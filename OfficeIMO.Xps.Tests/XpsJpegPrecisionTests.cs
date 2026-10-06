using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsJpegPrecisionTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void LosslessJpegPrecisionsRetainPaintThroughPortableSvgAndPdf(XpsFormat format) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Fixtures", "JpegLosslessPrecision");
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string name = row.Split(',')[0];
            byte[] jpeg = File.ReadAllBytes(Path.Combine(corpus, name));
            Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var source));
            var document = XpsDocument.Create(format);
            string uri = document.AddResource("Images/source.jpg", jpeg, "image/jpeg");
            document.AddPage(51, 33).AddImage(uri, 0, 0, 51, 33);
            var loaded = XpsDocument.Load(document.Save());
            var page = loaded.Pages[0];
            var svg = page.ToSvg();
            Assert.Empty(svg.Diagnostics);
            string href = XDocument.Parse(svg.Svg).Descendants().Single(e => e.Name.LocalName == "image").Attribute("href")!.Value;
            // Portable SVG receives decoded pixels rather than asking its consumer
            // to implement the uncommon JPEG process and sample precision.
            Assert.StartsWith("data:image/png;base64,", href);
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var svgDrawing));
            var images = new[] {
                OfficeDrawingRasterRenderer.Render(page.ToDrawing()),
                OfficeDrawingRasterRenderer.Render(svgDrawing!),
                OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(loaded.ToPdf()).Pages[0].ToDrawing(), scale: 4D / 3D)
            };
            for (int y = 0; y < 11; y++) for (int x = 0; x < 17; x++) {
                var expected = source!.GetPixel(x, y);
                foreach (var image in images) {
                    var actual = image.GetPixel(x * 3 + 1, y * 3 + 1);
                    Assert.True(Math.Abs(actual.R - expected.R) <= 1 && Math.Abs(actual.G - expected.G) <= 1 &&
                        Math.Abs(actual.B - expected.B) <= 1 && actual.A == expected.A, $"{name} at {x},{y}");
                }
            }
        }
    }
}
