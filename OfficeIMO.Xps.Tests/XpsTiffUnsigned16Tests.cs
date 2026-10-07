using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsTiffUnsigned16Tests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "Fixtures", "TiffUnsigned16");

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void IndependentlyEncodedSixteenBitTiffsRetainPaintAcrossRasterSvgAndPdf(XpsFormat format) {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            var document = XpsDocument.Create(format);
            string source = document.AddResource("Images/source.tif", File.ReadAllBytes(Path.Combine(Corpus, fields[0])), "image/tiff");
            document.AddPage(57, 39).AddImage(source, 0, 0, 57, 39);
            var page = XpsDocument.Load(document.Save()).Pages[0];
            if (fields[3] == "5" && fields[5] == "0") {
                Assert.Contains("CMYK image requires a usable ICC profile", page.ToSvg(true).Diagnostics);
                Assert.Throws<NotSupportedException>(() => page.ToDrawing());
                continue;
            }
            var svg = page.ToSvg();
            Assert.Empty(svg.Diagnostics);
            var direct = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var svgDrawing));
            var svgImage = OfficeDrawingRasterRenderer.Render(svgDrawing!, scale: 1, background: OfficeColor.White);
            var pdfPage = Assert.Single(PdfReadDocument.Open(document.ToPdf()).Pages);
            var pdfImage = OfficeDrawingRasterRenderer.Render(pdfPage.ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
            byte[] expected = File.ReadAllBytes(Path.Combine(Corpus, Path.ChangeExtension(fields[0], ".rgba")));
            for (int y = 0; y < 13; y++) for (int x = 0; x < 19; x++) {
                int offset = (y * 19 + x) * 4;
                foreach (var (image, lane) in new[] { (direct, "direct"), (svgImage, "svg"), (pdfImage, "pdf") }) {
                    var actual = image.GetPixel(x * 3 + 1, y * 3 + 1);
                    int Composite(int channel) => (expected[offset + channel] * expected[offset + 3] +
                        255 * (255 - expected[offset + 3]) + 127) / 255;
                    Assert.True(Math.Abs(actual.R - Composite(0)) <= 3, $"{fields[0]} {lane} {x},{y}: R {actual.R}, expected {Composite(0)}; dimensions {image.Width}x{image.Height}");
                    Assert.InRange(Math.Abs(actual.G - Composite(1)), 0, 3);
                    Assert.InRange(Math.Abs(actual.B - Composite(2)), 0, 3);
                    Assert.Equal(255, actual.A);
                }
            }
        }
    }
}
