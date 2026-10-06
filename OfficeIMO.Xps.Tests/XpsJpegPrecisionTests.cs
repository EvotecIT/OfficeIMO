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
    [InlineData(XpsFormat.Xps, "TiffJpegLosslessPrecision", true)]
    [InlineData(XpsFormat.OpenXps, "TiffJpegLosslessPrecision", true)]
    [InlineData(XpsFormat.Xps, "TiffJpegArithmeticLossless", true)]
    [InlineData(XpsFormat.OpenXps, "TiffJpegArithmeticLossless", true)]
    [InlineData(XpsFormat.Xps, "JpegArithmeticLossless")]
    [InlineData(XpsFormat.OpenXps, "JpegArithmeticLossless")]
    [InlineData(XpsFormat.Xps, "JpegArithmeticProgressive")]
    [InlineData(XpsFormat.OpenXps, "JpegArithmeticProgressive")]
    [InlineData(XpsFormat.Xps, "JpegArithmetic", true)]
    [InlineData(XpsFormat.OpenXps, "JpegArithmetic", true)]
    [InlineData(XpsFormat.Xps, "JpegArithmetic")]
    [InlineData(XpsFormat.OpenXps, "JpegArithmetic")]
    [InlineData(XpsFormat.Xps, "JpegLosslessPrecision")]
    [InlineData(XpsFormat.Xps, "JpegDct12")]
    [InlineData(XpsFormat.OpenXps, "JpegLosslessPrecision")]
    [InlineData(XpsFormat.OpenXps, "JpegDct12")]
    public void JpegPrecisionsRetainPaintThroughPortableSvgAndPdf(XpsFormat format, string folder, bool tiff = false) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Fixtures", folder);
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string name = row.Split(',')[0];
            byte[] jpeg = File.ReadAllBytes(Path.Combine(corpus, name + (tiff && !name.EndsWith(".tif", StringComparison.OrdinalIgnoreCase) ? ".tif" : "")));
            OfficeRasterImage? source;
            Assert.True(tiff ? OfficeTiffCodec.TryDecode(jpeg, out source) : OfficeJpegCodec.TryDecode(jpeg, out source));
            var document = XpsDocument.Create(format);
            string uri = document.AddResource(tiff ? "Images/source.tif" : "Images/source.jpg", jpeg, tiff ? "image/tiff" : "image/jpeg");
            int width = source!.Width, height = source.Height;
            document.AddPage(width * 3, height * 3).AddImage(uri, 0, 0, width * 3, height * 3);
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
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
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
