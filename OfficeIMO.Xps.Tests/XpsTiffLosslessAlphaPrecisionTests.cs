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
    [InlineData(XpsFormat.Xps, false, "TiffJpegChromaAlpha")]
    [InlineData(XpsFormat.OpenXps, false, "TiffJpegChromaAlpha")]
    [InlineData(XpsFormat.Xps, true, "TiffJpegChromaAlpha")]
    [InlineData(XpsFormat.OpenXps, true, "TiffJpegChromaAlpha")]
    [InlineData(XpsFormat.Xps, false, "TiffJpegArithmeticAlphaPrecision")]
    [InlineData(XpsFormat.OpenXps, false, "TiffJpegArithmeticAlphaPrecision")]
    [InlineData(XpsFormat.Xps, true, "TiffJpegArithmeticAlphaPrecision")]
    [InlineData(XpsFormat.OpenXps, true, "TiffJpegArithmeticAlphaPrecision")]
    [InlineData(XpsFormat.Xps, false, "TiffJpegArithmetic4x4")]
    [InlineData(XpsFormat.OpenXps, false, "TiffJpegArithmetic4x4")]
    [InlineData(XpsFormat.Xps, true, "TiffJpegArithmetic4x4")]
    [InlineData(XpsFormat.OpenXps, true, "TiffJpegArithmetic4x4")]
    public void NativePrecisionAlphaRetainsVisiblePaint(XpsFormat format, bool profiled, string corpusName = "TiffJpegLosslessAlphaPrecision") {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Fixtures", corpusName);
        foreach (string line in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] row = line.Split(',');
            if (profiled && row[1] is not ("2" or "5" or "6")) continue;
            if (!profiled && row[1] == "5") continue; // XPS requires an explicit CMYK profile.
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, row[0] + (profiled ? ".icc-rgba" : ".rgba")));
            int width = corpusName != "TiffJpegLosslessAlphaPrecision" ? int.Parse(row[5]) : 35;
            int height = corpusName != "TiffJpegLosslessAlphaPrecision" ? int.Parse(row[6]) : 19;
            var document = XpsDocument.Create(format);
            string resource = document.AddResource("Images/source.tif", File.ReadAllBytes(Path.Combine(corpus, row[0])), "image/tiff");
            document.AddPage(width * 3, height * 3).AddImage(resource, 0, 0, width * 3, height * 3);
            if (profiled) {
                string profile = document.AddResource("Profiles/source.icc", File.ReadAllBytes(Path.Combine(corpus, "..", "IccColorCorpus", row[1] == "5" ? "littlecms-cmyk-lut.icc" : "icc-dci-p3-matrix.icc")), "application/vnd.ms-color.iccprofile");
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
                for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                    int at = (y * width + x) * 4, alpha = expected[at + 3], tolerance = profiled ? 3 : 2;
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
