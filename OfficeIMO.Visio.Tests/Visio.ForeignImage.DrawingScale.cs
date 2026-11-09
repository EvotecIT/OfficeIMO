using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public partial class VisioForeignImageRenderingTests {
    [Fact]
    public void ForeignCropAndMirroredPixelsFollowPhysicalPageScaleAfterReopening() {
        using var input = Fixture();
        XDocument xml = XDocument.Load(input);
        XElement foreign = xml.Descendants(Ns + "Foreign").Single();
        foreign.Element(Ns + "ImgOffsetX")!.Value = "-1";
        foreign.Element(Ns + "ImgWidth")!.Value = "4";
        xml.Descendants(Ns + "XForm").Single().Add(new XElement(Ns + "FlipX", "1"));
        VisioDocument document = Load(xml);
        OfficeRasterImage expectedPixels = VisualBaselineTestSupport.DecodePng(document.Pages[0].ToImage(new VisioImageExportOptions {
            TargetDpi = 24, Supersampling = 1, RenderText = false
        }).Bytes, "Reference foreign PNG did not decode.");
        XElement expectedImage = XDocument.Parse(document.Pages[0].ToSvg(new VisioSvgSaveOptions {
            PixelsPerInch = 24, RenderText = false
        })).Descendants().Single(node => node.Name.LocalName == "image");
        document.Pages[0].PageScale = new VisioScaleSetting(0.5, VisioMeasurementUnit.Inches);
        foreach (VisioDocument candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            byte[] before = candidate.ToLegacyXmlResult().Value;
            VisioPage page = candidate.Pages[0];
            byte[] actualPixels = page.ToImage(new VisioImageExportOptions {
                TargetDpi = 48, Supersampling = 1, RenderText = false
            }).Bytes;
            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(actualPixels, "Scaled foreign PNG did not decode.");
            Assert.Equal(96, image.Width);
            Assert.Equal(96, image.Height);
            Assert.Equal(expectedPixels.GetPixels(), image.GetPixels());
            XElement actualImage = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions {
                PixelsPerInch = 48, RenderText = false
            })).Descendants().Single(node => node.Name.LocalName == "image");
            Assert.Equal(expectedImage.ToString(), actualImage.ToString());
            Assert.Equal(before, candidate.ToLegacyXmlResult().Value);
        }
    }
}
