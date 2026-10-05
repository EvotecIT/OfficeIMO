using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsTiffFaxTests {
    [Theory]
    [InlineData(XpsFormat.Xps, "TiffFax", 83, 19)]
    [InlineData(XpsFormat.Xps, "TiffFaxUncompressed", 32, 12)]
    [InlineData(XpsFormat.OpenXps, "TiffFax", 83, 19)]
    [InlineData(XpsFormat.OpenXps, "TiffFaxUncompressed", 32, 12)]
    public void FaxTiffResourcesRetainRasterSvgAndPdfPaint(XpsFormat format, string folder, int width, int height) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Fixtures", folder);
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            var document = XpsDocument.Create(format);
            string resource = document.AddResource("Images/packed.tif", File.ReadAllBytes(Path.Combine(corpus, fields[0])), "image/tiff");
            document.AddPage(width * 3, height * 3).AddImage(resource, 0, 0, width * 3, height * 3);
            var page = XpsDocument.Load(document.Save()).Pages[0];
            var svg = page.ToSvg();
            Assert.Empty(svg.Diagnostics);
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var drawing));
            var native = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
            var svgImage = OfficeDrawingRasterRenderer.Render(drawing!, background: OfficeColor.White);
            var pdfImage = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(document.ToPdf()).Pages[0].ToDrawing(),
                scale: 4D / 3D, background: OfficeColor.White);
            byte[]? reference = folder == "TiffFaxUncompressed" ? File.ReadAllBytes(Path.Combine(corpus, fields[0] + ".rgb")) : null;
            bool whiteIsZero = fields[3] == "0";
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                int value = y % 7 == 0 ? 0 : y % 7 == 1 ? 1 : y % 7 == 2 ? x & 1 : (x / 13 + y / 2) & 1;
                int red = reference != null ? reference[(y * width + x) * 3] : (whiteIsZero ? 1 - value : value) * 255, green = red, blue = red;
                foreach (var image in new[] { native, svgImage, pdfImage }) {
                    var actual = image.GetPixel(x * 3 + 1, y * 3 + 1);
                    Assert.True(Math.Abs(actual.R - red) <= 1 && Math.Abs(actual.G - green) <= 1 &&
                        Math.Abs(actual.B - blue) <= 1 && actual.A == 255, $"{fields[0]} at {x},{y}");
                }
            }
        }
    }
}
