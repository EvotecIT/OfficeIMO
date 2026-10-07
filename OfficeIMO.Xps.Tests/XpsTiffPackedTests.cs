using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsTiffPackedTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void PackedTiffResourcesRetainRasterSvgAndPdfPaint(XpsFormat format) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Fixtures", "TiffPacked");
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            var document = XpsDocument.Create(format);
            string resource = document.AddResource("Images/packed.tif", File.ReadAllBytes(Path.Combine(corpus, fields[0])), "image/tiff");
            document.AddPage(57, 51).AddImage(resource, 0, 0, 57, 51);
            var page = XpsDocument.Load(document.Save()).Pages[0];
            var svg = page.ToSvg();
            Assert.Empty(svg.Diagnostics);
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var drawing));
            var native = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
            var svgImage = OfficeDrawingRasterRenderer.Render(drawing!, background: OfficeColor.White);
            var pdfImage = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(document.ToPdf()).Pages[0].ToDrawing(),
                scale: 4D / 3D, background: OfficeColor.White);
            int bits = int.Parse(fields[1]), photo = int.Parse(fields[2]), mask = (1 << bits) - 1;
            for (int y = 0; y < 17; y++) for (int x = 0; x < 19; x++) {
                int value = (x * 3 + y * 5) & mask;
                int red = (photo == 0 ? mask - value : value) * 255 / mask;
                int green = photo == 3 ? (mask - value) * 255 / mask : red;
                int blue = photo == 3 ? (value * 7 & mask) * 255 / mask : red;
                foreach (var image in new[] { native, svgImage, pdfImage }) {
                    var actual = image.GetPixel(x * 3 + 1, y * 3 + 1);
                    Assert.True(Math.Abs(actual.R - red) <= 1 && Math.Abs(actual.G - green) <= 1 &&
                        Math.Abs(actual.B - blue) <= 1 && actual.A == 255, $"{fields[0]} at {x},{y}");
                }
            }
        }
    }
}
