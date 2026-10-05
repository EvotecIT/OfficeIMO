using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsTiffJpegTests {
    [Theory]
    [InlineData(XpsFormat.Xps, "TiffJpegAlpha", 35, 19)]
    [InlineData(XpsFormat.OpenXps, "TiffJpegAlpha", 35, 19)]
    [InlineData(XpsFormat.Xps, "TiffJpegExtended", 35, 19)]
    [InlineData(XpsFormat.OpenXps, "TiffJpegExtended", 35, 19)]
    [InlineData(XpsFormat.Xps, "TiffJpeg", 35, 19)]
    [InlineData(XpsFormat.OpenXps, "TiffJpeg", 35, 19)]
    [InlineData(XpsFormat.Xps, "TiffJpegPlanar", 67, 35)]
    [InlineData(XpsFormat.OpenXps, "TiffJpegPlanar", 67, 35)]
    [InlineData(XpsFormat.Xps, "TiffJpegCosited", 0, 0)]
    [InlineData(XpsFormat.OpenXps, "TiffJpegCosited", 0, 0)]
    public void JpegTiffResourcesRetainRasterSvgAndPdfPaint(XpsFormat format, string folder, int width, int height) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Fixtures", folder);
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            if (folder == "TiffJpegCosited") { width = int.Parse(fields[7]); height = int.Parse(fields[8]); }
            var document = XpsDocument.Create(format);
            string resource = document.AddResource("Images/source.tif", File.ReadAllBytes(Path.Combine(corpus, fields[0])), "image/tiff");
            document.AddPage(width * 3, height * 3).AddImage(resource, 0, 0, width * 3, height * 3);
            if ((folder == "TiffJpeg" || folder == "TiffJpegExtended" || folder == "TiffJpegAlpha") && fields[1] == "5") {
                var originalPage = document.Pages[0];
                Assert.Throws<NotSupportedException>(() => originalPage.ToDrawing());
                byte[] profile = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus", "littlecms-cmyk-lut.icc"));
                string profileUri = document.AddResource("Profiles/source.icc", profile, "application/vnd.ms-color.iccprofile");
                var markup = originalPage.GetMarkup();
                markup.Descendants().Single(e => e.Name.LocalName == "ImageBrush").SetAttributeValue("ImageSource",
                    "{ColorConvertedBitmap " + resource + " " + profileUri + "}");
                originalPage.ReplaceMarkup(markup);
            }
            var page = XpsDocument.Load(document.Save()).Pages[0];
            if (folder == "TiffJpegAlpha") {
                Assert.True(OfficeTiffCodec.TryDecode(File.ReadAllBytes(Path.Combine(corpus, fields[0])), out var source));
                var transparent = OfficeDrawingRasterRenderer.Render(page.ToDrawing());
                for (int ay = 0; ay < height; ay++) for (int ax = 0; ax < width; ax++)
                    Assert.Equal(source!.GetPixel(ax, ay).A, transparent.GetPixel(ax * 3 + 1, ay * 3 + 1).A);
            }
            var svg = page.ToSvg();
            Assert.Empty(svg.Diagnostics);
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var drawing));
            var native = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
            var svgImage = OfficeDrawingRasterRenderer.Render(drawing!, background: OfficeColor.White);
            var pdfImage = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(document.ToPdf()).Pages[0].ToDrawing(),
                scale: 4D / 3D, background: OfficeColor.White);
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                var expected = native.GetPixel(x * 3 + 1, y * 3 + 1);
                foreach (var image in new[] { svgImage, pdfImage }) {
                    var actual = image.GetPixel(x * 3 + 1, y * 3 + 1);
                    Assert.True(Math.Abs(actual.R - expected.R) <= 1 && Math.Abs(actual.G - expected.G) <= 1 &&
                        Math.Abs(actual.B - expected.B) <= 1 && actual.A == 255, $"{fields[0]} at {x},{y}");
                }
            }
        }
    }
}
