using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsJpegXrTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "Fixtures", "JpegXr");

    [Theory]
    [InlineData(false, "rgb-19x13-frequency-overlap2-alpha0", false)]
    [InlineData(true, "rgb-19x13-frequency-overlap2-alpha0", false)]
    [InlineData(false, "u16-3c-spatial-overlap0-q0-alpha0", false)]
    [InlineData(true, "u16-3c-spatial-overlap0-q0-alpha0", false)]
    [InlineData(false, "u16-1c-spatial-overlap1-q0-alpha0", true)]
    [InlineData(true, "u16-1c-spatial-overlap1-q0-alpha0", true)]
    [InlineData(false, "u16-premultiplied-alpha1", false)]
    [InlineData(true, "u16-premultiplied-alpha2", false)]
    public void EmbeddedOrAssociatedProfileUsesTheExistingImageColorPipeline(bool associated, string id, bool gray) {
        byte[] encoded = File.ReadAllBytes(Path.Combine(Corpus, id + ".jxr"));
        byte[] profile = File.ReadAllBytes(gray
            ? Path.Combine(AppContext.BaseDirectory, "Fixtures", "ColorImages", "gray-gamma18.icc")
            : Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus", "littlecms-rgb-matrix.icc"));
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out var raw));
        byte[] png = OfficeRasterImageEncoder.Encode(raw!, OfficeImageExportFormat.Png);
        (OfficeRasterImage Raster, byte[] Normalized) Paint(byte[] bytes, string type, bool embed) {
            if (embed) bytes = OfficeIMO.TestAssets.JpegXrTestFixture.WithField(bytes, 0x8773, 7, profile);
            var document = XpsDocument.Create();
            string uri = document.AddResource("Images/source", bytes, type);
            var page = document.AddPage(57, 39).AddImage(uri, 0, 0, 57, 39);
            if (!embed) {
                string icc = document.AddResource("Profiles/source.icc", profile, "application/vnd.ms-color.iccprofile");
                var markup = page.GetMarkup();
                markup.Descendants().Single(e => e.Name.LocalName == "ImageBrush").SetAttributeValue("ImageSource", "{ColorConvertedBitmap " + uri + " " + icc + "}");
                page.ReplaceMarkup(markup);
            }
            page = XpsDocument.Load(document.Save()).Pages[0];
            var svg = XDocument.Parse(page.ToSvg().Svg);
            string data = (string)svg.Descendants().Single(e => e.Name.LocalName == "image").Attribute("href")!;
            Assert.True(OfficeRasterImageDecoder.TryDecode(Convert.FromBase64String(data.Substring(data.IndexOf(',') + 1)), out var normalized));
            return (OfficeDrawingRasterRenderer.Render(page.ToDrawing()), normalized!.GetPixels());
        }
        bool sixteenBit = id.StartsWith("u16-", StringComparison.Ordinal);
        var expected = Paint(sixteenBit ? File.ReadAllBytes(Path.Combine(Corpus, id + ".tif")) : png,
            sixteenBit ? "image/tiff" : "image/png", false);
        var actual = Paint(encoded, "image/jxr", !associated);
        Assert.Equal(expected.Normalized, actual.Normalized);
        byte[] expectedPaint = expected.Raster.GetPixels(), actualPaint = actual.Raster.GetPixels();
        Assert.Equal(expectedPaint.Length, actualPaint.Length);
        for (int i = 0; i < expectedPaint.Length; i++) Assert.InRange(Math.Abs(expectedPaint[i] - actualPaint[i]), 0, 1);
    }

    [Theory]
    [InlineData(XpsFormat.Xps, "image/vnd.ms-photo")]
    [InlineData(XpsFormat.OpenXps, "image/jxr")]
    public void JpegXrBrushPixelsSurviveSvgRasterAndPdfReader(XpsFormat format, string contentType) {
        foreach (string sourceName in new[] {
            "rgb-19x13-spatial-overlap2-alpha0", "rgb-19x13-frequency-overlap1-alpha0",
            "l-19x13-frequency-overlap2-alpha0", "rgba-19x13-spatial-overlap1-alpha1",
            "rgba-19x13-frequency-overlap2-alpha2",
            "u16-1c-spatial-overlap1-q0-alpha0", "u16-3c-spatial-overlap0-q0-alpha0",
            "u16-4c-frequency-overlap2-q0-alpha1", "u16-4c-frequency-overlap2-q0-alpha2",
            "u16-premultiplied-alpha1", "u16-premultiplied-alpha2",
            "subsampled-yuv420-spatial-overlap0-q32-alpha0",
            "subsampled-yuv420-frequency-overlap1-q32-alpha1",
            "subsampled-yuv422-spatial-overlap2-q32-alpha2",
            "subsampled-yuv422-frequency-overlap0-q32-alpha0",
            "extended-s16-3c-frequency-q32-alpha0", "extended-s32-3c-frequency-q32-alpha0",
            "extended-f16-3c-frequency-q32-alpha0", "extended-f32-3c-frequency-q32-alpha0",
            "extended-f32-premultiplied-4c-frequency-q32-alpha2", "extended-s32-limits-3c-frequency-q32-alpha0",
            "mixed-8-spatial-p0-a1", "mixed-8-frequency-p0-a2", "mixed-8-frequency-p0-a3",
            "mixed-16-spatial-p1-a2", "mixed-16-frequency-p1-a3", "mixed-16-frequency-p2-a3"
        }) {
            byte[] encoded = File.ReadAllBytes(Path.Combine(Corpus, sourceName + ".jxr"));
            byte[] expected = File.ReadAllBytes(Path.Combine(Corpus, sourceName + ".rgba"));
            var document = XpsDocument.Create(format);
            string uri = document.AddResource("Images/source.jxr", encoded, contentType);
            Assert.True(OfficeImageReader.TryIdentifyByContent(encoded, null, out var metadata));
            int width = metadata.Width, height = metadata.Height;
            document.AddPage(width * 3, height * 3).AddImage(uri, 0, 0, width * 3, height * 3);
            var page = XpsDocument.Load(document.Save()).Pages[0];
            var svg = page.ToSvg(); Assert.Empty(svg.Diagnostics); Assert.Contains("data:image/png;base64,", svg.Svg);
            var direct = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var drawing));
            var svgRaster = OfficeDrawingRasterRenderer.Render(drawing!, background: OfficeColor.White);
            var pdfPage = Assert.Single(PdfReadDocument.Open(document.ToPdf()).Pages);
            var pdfRaster = OfficeDrawingRasterRenderer.Render(pdfPage.ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                int offset = (y * width + x) * 4;
                foreach (var raster in new[] { direct, svgRaster, pdfRaster }) {
                    var pixel = raster.GetPixel(x * 3 + 1, y * 3 + 1);
                    int Composite(int c) => (expected[offset + c] * expected[offset + 3] + 255 * (255 - expected[offset + 3]) + 127) / 255;
                    Assert.InRange(Math.Abs(pixel.R - Composite(0)), 0, 1);
                    Assert.InRange(Math.Abs(pixel.G - Composite(1)), 0, 1);
                    Assert.InRange(Math.Abs(pixel.B - Composite(2)), 0, 1);
                    Assert.Equal(255, pixel.A);
                }
            }
        }
    }
}
