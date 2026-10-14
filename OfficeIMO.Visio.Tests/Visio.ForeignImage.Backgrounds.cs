using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public partial class VisioForeignImageRenderingTests {
    [Fact]
    public void BackgroundCropMirroringAndPhysicalScaleAgreeAcrossRasterSvgAndDrawingAfterReopening() {
        using var input = Fixture();
        XDocument xml = XDocument.Load(input);
        XElement foreign = xml.Descendants(Ns + "Foreign").Single();
        foreign.Element(Ns + "ImgOffsetX")!.Value = "-1";
        foreign.Element(Ns + "ImgWidth")!.Value = "4";
        xml.Descendants(Ns + "XForm").Single().Add(new XElement(Ns + "FlipX", "1"));
        VisioDocument source = Load(xml);
        VisioPage background = source.Pages[0];
        background.PageScale = new VisioScaleSetting(.5, VisioMeasurementUnit.Inches);
        source.AddPage("Foreground", 4, 3).SetBackgroundPage(background);
        foreach (VisioDocument document in new[] { source,
            VisioDocument.LoadLegacyXml(new MemoryStream(source.ToLegacyXmlResult().Value)).Value,
            VisioDocument.Load(new MemoryStream(source.ToBytes())) }) {
            byte[] before = document.ToLegacyXmlResult().Value;
            var options = new VisioPngSaveOptions { PixelsPerInch = 72, Supersampling = 1, RenderText = false, RenderStencilArtwork = false };
            OfficeRasterImage reference = DecodeBackground(document.Pages[0].ToPng(options));
            OfficeRasterImage native = DecodeBackground(document.Pages[1].ToPng(options));
            OfficeRasterImage drawing = DecodeBackground(OfficeDrawingRasterRenderer.ToPng(document.Pages[1].ToDrawing().Value, background: OfficeColor.White));
            byte[] svgBytes = Encoding.UTF8.GetBytes(document.Pages[1].ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 72, RenderText = false }));
            Assert.True(VisioSvgPreviewRasterizer.TryRasterize(svgBytes, null, null, null, null, null, null, null, default, out OfficeRasterImage? svg));
            foreach ((int x, int y) in new[] { (54, 54), (90, 54), (54, 90), (90, 90) }) {
                Assert.Equal(reference.GetPixel(x, y), native.GetPixel(x, y + 72));
                Assert.Equal(reference.GetPixel(x, y), drawing.GetPixel(x, y + 72));
                Assert.Equal(reference.GetPixel(x, y), svg!.GetPixel(x, y + 72));
            }
            Assert.Throws<InvalidDataException>(() => document.ToDrawings(new VisioDrawingOptions { MaximumTotalImagePixels = 16 * 16 }));
            Assert.Equal(before, document.ToLegacyXmlResult().Value);
        }
    }

    private static OfficeRasterImage DecodeBackground(byte[] bytes) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? image)); return image!;
    }
}
