using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsTiffLosslessAlphaPrecisionTests {
    [Theory]
    [InlineData(XpsFormat.Xps, false)]
    [InlineData(XpsFormat.OpenXps, false)]
    [InlineData(XpsFormat.Xps, true)]
    [InlineData(XpsFormat.OpenXps, true)]
    public void NativePrecisionAlphaRetainsVisiblePaint(XpsFormat format, bool profiled) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Fixtures", "TiffJpegLosslessAlphaPrecision");
        foreach (string line in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] row = line.Split(',');
            if (profiled && row[1] is not ("2" or "6")) continue;
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, row[0] + (profiled ? ".icc-rgba" : ".rgba")));
            var document = XpsDocument.Create(format);
            string resource = document.AddResource("Images/source.tif", File.ReadAllBytes(Path.Combine(corpus, row[0])), "image/tiff");
            document.AddPage(105, 57).AddImage(resource, 0, 0, 105, 57);
            if (profiled) {
                string profile = document.AddResource("Profiles/source.icc", File.ReadAllBytes(Path.Combine(corpus, "..", "IccColorCorpus", "icc-dci-p3-matrix.icc")), "application/vnd.ms-color.iccprofile");
                var markup = document.Pages[0].GetMarkup();
                markup.Descendants().Single(e => e.Name.LocalName == "ImageBrush").SetAttributeValue("ImageSource",
                    "{ColorConvertedBitmap " + resource + " " + profile + "}");
                document.Pages[0].ReplaceMarkup(markup);
            }
            var loaded = XpsDocument.Load(document.Save());
            var page = loaded.Pages[0];
            var nativeDrawing = page.ToDrawing();
            var transparent = OfficeDrawingRasterRenderer.Render(nativeDrawing);
            var svg = page.ToSvg();
            Assert.Empty(svg.Diagnostics);
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var svgDrawing));
            var pdfDrawing = PdfReadDocument.Open(loaded.ToPdf()).Pages[0].ToDrawing();
            foreach (var background in new[] { OfficeColor.White, OfficeColor.Black }) {
                var images = new[] {
                    OfficeDrawingRasterRenderer.Render(nativeDrawing, background: background),
                    OfficeDrawingRasterRenderer.Render(svgDrawing!, background: background),
                    OfficeDrawingRasterRenderer.Render(pdfDrawing, scale: 4D / 3D, background: background)
                };
                for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                    int at = (y * 35 + x) * 4, alpha = expected[at + 3], tolerance = profiled ? 3 : 2;
                    Assert.Equal(alpha, transparent.GetPixel(x * 3 + 1, y * 3 + 1).A);
                    int Composite(int c) => (expected[at + c] * alpha + background.R * (255 - alpha) + 127) / 255;
                    foreach (var image in images) {
                        var actual = image.GetPixel(x * 3 + 1, y * 3 + 1);
                        Assert.True(Math.Abs(actual.R - Composite(0)) <= tolerance && Math.Abs(actual.G - Composite(1)) <= tolerance &&
                            Math.Abs(actual.B - Composite(2)) <= tolerance && actual.A == 255, $"{row[0]} {format} profile={profiled} at {x},{y}");
                    }
                }
            }
        }
    }
}
