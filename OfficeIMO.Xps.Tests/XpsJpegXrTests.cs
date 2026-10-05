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
    [InlineData(false)]
    [InlineData(true)]
    public void EmbeddedOrAssociatedProfileUsesTheExistingImageColorPipeline(bool associated) {
        string id = "rgb-19x13-frequency-overlap2-alpha0";
        byte[] encoded = File.ReadAllBytes(Path.Combine(Corpus, id + ".jxr"));
        byte[] profile = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus", "littlecms-rgb-matrix.icc"));
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
        var expected = Paint(png, "image/png", false);
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
            "rgba-19x13-frequency-overlap2-alpha2"
        }) {
            byte[] encoded = File.ReadAllBytes(Path.Combine(Corpus, sourceName + ".jxr"));
            byte[] expected = File.ReadAllBytes(Path.Combine(Corpus, sourceName + ".rgba"));
            var document = XpsDocument.Create(format);
            string uri = document.AddResource("Images/source.jxr", encoded, contentType);
            document.AddPage(57, 39).AddImage(uri, 0, 0, 57, 39);
            var page = XpsDocument.Load(document.Save()).Pages[0];
            var svg = page.ToSvg(); Assert.Empty(svg.Diagnostics); Assert.Contains("data:image/png;base64,", svg.Svg);
            var direct = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var drawing));
            var svgRaster = OfficeDrawingRasterRenderer.Render(drawing!, background: OfficeColor.White);
            var pdfPage = Assert.Single(PdfReadDocument.Open(document.ToPdf()).Pages);
            var pdfRaster = OfficeDrawingRasterRenderer.Render(pdfPage.ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
            for (int y = 0; y < 13; y++) for (int x = 0; x < 19; x++) {
                int offset = (y * 19 + x) * 4;
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
