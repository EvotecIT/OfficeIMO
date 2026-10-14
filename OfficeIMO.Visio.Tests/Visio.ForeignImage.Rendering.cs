using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public partial class VisioForeignImageRenderingTests {
    private static readonly XNamespace Ns = "http://schemas.microsoft.com/visio/2003/core";

    [Fact]
    public void LegacyBitmapRendersInSvgAndRasterAfterPackageRoundTrip() {
        using var input = Fixture();
        var document = VisioDocument.LoadLegacyXml(input).Value;
        foreach (var candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            var page = candidate.Pages[0];
            string svg = page.ToSvg(new VisioSvgSaveOptions { RenderText = false, RenderStencilArtwork = false });
            Assert.Contains("data-officeimo-foreign-image", svg);
            var imageElement = XDocument.Parse(svg).Descendants().Single(element => element.Name.LocalName == "image");
            string data = imageElement.Attributes().Single(attribute => attribute.Name.LocalName == "href").Value;
            Assert.True(OfficeRasterImageDecoder.TryDecode(Convert.FromBase64String(data.Substring(data.IndexOf(',') + 1)), out var embedded));
            Assert.Equal(OfficeColor.Red, embedded!.GetPixel(0, 0));
            var raster = Render(page);
            Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 25));
            Assert.Equal(OfficeColor.Green, raster.GetPixel(55, 25));
            Assert.Equal(OfficeColor.Blue, raster.GetPixel(25, 55));
            Assert.Equal(OfficeColor.White, raster.GetPixel(55, 55));
        }
    }

    [Fact]
    public void ImageCropAndFormulaDimensionsFollowTypedResize() {
        using var input = Fixture(); var xml = XDocument.Load(input);
        var foreign = xml.Descendants(Ns + "Foreign").Single();
        foreign.Element(Ns + "ImgOffsetX")!.Value = "-1";
        foreign.Element(Ns + "ImgWidth")!.SetAttributeValue("F", "Width*2");
        foreign.Element(Ns + "ImgWidth")!.Value = "4";
        var document = Load(xml); var shape = document.Pages[0].Shapes[0];
        var raster = Render(document.Pages[0]);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 25));
        Assert.Equal(OfficeColor.Green, raster.GetPixel(55, 25));
        shape.Width = 1;
        var diagnostics = new System.Collections.Generic.List<OfficeImageExportDiagnostic>();
        Assert.True(VisioForeignImage.TryGetProjection(shape, document.Pages[0], 20, diagnostics, null, out var projection));
        Assert.Equal(0.5, projection.SourceCrop.Left);
        Assert.Equal(0, projection.SourceCrop.Right);
        Assert.Equal(20, projection.Width);
        Assert.Empty(diagnostics);
        Assert.Equal(OfficeColor.Green, Render(document.Pages[0]).GetPixel(25, 25));
    }

    [Fact]
    public void ExplicitResizeRetainsImageCropFractionsPayloadAndPixelsAcrossReopening() {
        using var input = Fixture(); var xml = XDocument.Load(input);
        XElement foreign = xml.Descendants(Ns + "Foreign").Single();
        foreign.Element(Ns + "ImgOffsetX")!.Value = "-1";
        foreign.Element(Ns + "ImgWidth")!.SetAttributeValue("F", "Width*2");
        foreign.Element(Ns + "ImgWidth")!.Value = "4";
        var document = Load(xml); var page = document.Pages[0]; var shape = page.Shapes[0];
        byte[] original = document.ToLegacyXmlResult().Value;
        page.ResizeShape(shape, 4, 1);
        Assert.True(VisioForeignImage.TryGetProjection(shape, page, 20, null, null, out var projection));
        Assert.Equal(.25, projection.SourceCrop.Left); Assert.Equal(.25, projection.SourceCrop.Right);
        foreach (var reopened in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())),
            VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value }) {
            Assert.True(VisioForeignImage.TryGetProjection(reopened.Pages[0].Shapes[0], reopened.Pages[0], 20, null, null, out var saved));
            Assert.Equal(projection.X, saved.X, 8); Assert.Equal(projection.Y, saved.Y, 8);
            Assert.Equal(projection.Width, saved.Width, 8); Assert.Equal(projection.Height, saved.Height, 8);
            Assert.Equal(projection.SourceCrop, saved.SourceCrop);
            XElement sourcePayload = XDocument.Load(new MemoryStream(original)).Descendants(Ns + "ForeignData").Single();
            XElement savedPayload = XDocument.Load(new MemoryStream(reopened.ToLegacyXmlResult().Value)).Descendants(Ns + "ForeignData").Single();
            Assert.Equal(sourcePayload.Value, savedPayload.Value);
            Assert.Equal(page.ToPng(new VisioPngSaveOptions { PixelsPerInch = 20, Supersampling = 1, RenderText = false, RenderStencilArtwork = false }),
                reopened.Pages[0].ToPng(new VisioPngSaveOptions { PixelsPerInch = 20, Supersampling = 1, RenderText = false, RenderStencilArtwork = false }));
        }
    }

    [Fact]
    public void GroupRotationAndMirroringApplyToEmbeddedPixels() {
        using var input = Fixture(); var xml = XDocument.Load(input);
        var shape = xml.Descendants(Ns + "Shape").Single();
        shape.Element(Ns + "XForm")!.Add(new XElement(Ns + "FlipX", "1"));
        var group = new XElement(Ns + "Shape", new XAttribute("ID", "2"), new XAttribute("Type", "Group"),
            new XElement(Ns + "XForm", new XElement(Ns + "PinX", "2"), new XElement(Ns + "PinY", "2"),
                new XElement(Ns + "LocPinX", "2"), new XElement(Ns + "LocPinY", "2"),
                new XElement(Ns + "Width", "4"), new XElement(Ns + "Height", "4"), new XElement(Ns + "Angle", Math.PI / 2)),
            new XElement(Ns + "Line", new XElement(Ns + "LinePattern", "0")),
            new XElement(Ns + "Fill", new XElement(Ns + "FillPattern", "0")));
        shape.ReplaceWith(group); group.Add(new XElement(Ns + "Shapes", shape));
        var raster = Render(Load(xml).Pages[0]);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 25));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(55, 25));
        Assert.Equal(OfficeColor.Green, raster.GetPixel(25, 55));
        Assert.Equal(OfficeColor.White, raster.GetPixel(55, 55));
    }

    [Fact]
    public void UnsupportedForeignObjectsProduceVisibleFallbackAndLossDiagnostics() {
        using var input = Fixture(); var xml = XDocument.Load(input);
        xml.Descendants(Ns + "ForeignData").Single().SetAttributeValue("ForeignType", "Object");
        var document = Load(xml);
        var codec = new RecordingCodec();
        var result = document.ExportImage(OfficeImageExportFormat.Svg, new VisioImageExportOptions { RenderText = false, ImageCodec = codec });
        Assert.Equal(0, codec.Calls);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains("data-officeimo-foreign-image", Encoding.UTF8.GetString(result.Bytes));
        // Even recognizable image bytes labelled OLE remain opaque, including to caller codecs.
        Assert.Equal(OfficeColor.FromRgb(245, 245, 245), Render(document.Pages[0]).GetPixel(30, 25));
    }

    [Fact]
    public void UnsupportedPlacementFormulaUsesCacheWithDiagnosticAndInvalidValuesAreSkipped() {
        using var input = Fixture(); var xml = XDocument.Load(input);
        var width = xml.Descendants(Ns + "ImgWidth").Single(); width.SetAttributeValue("F", "Sheet.99!Width");
        var result = Load(xml).ExportImage(OfficeImageExportFormat.Svg, new VisioImageExportOptions { RenderText = false });
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "VISIO_FOREIGN_CACHED_FORMULA");
        width.Value = "NaN";
        result = Load(xml).ExportImage(OfficeImageExportFormat.Svg, new VisioImageExportOptions { RenderText = false });
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "VISIO_FOREIGN_PLACEMENT" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.DoesNotContain("data-officeimo-foreign-image", Encoding.UTF8.GetString(result.Bytes));
    }

    [Fact]
    public void MasterImageAndInheritedPlacementFormulaFollowInstanceSize() {
        using var input = Fixture(); var xml = XDocument.Load(input);
        var instance = xml.Descendants(Ns + "Shape").Single();
        var masterShape = new XElement(instance);
        masterShape.Descendants(Ns + "ImgWidth").Single().SetAttributeValue("F", "Width*1");
        xml.Root!.AddFirst(new XElement(Ns + "Masters", new XElement(Ns + "Master", new XAttribute("ID", "0"),
            new XAttribute("NameU", "Picture"), new XElement(Ns + "Shapes", masterShape))));
        instance.SetAttributeValue("Master", "0");
        instance.Element(Ns + "ForeignData")!.Remove();
        var width = instance.Descendants(Ns + "ImgWidth").Single(); width.SetAttributeValue("F", "Inh"); width.Value = "99";
        var document = Load(xml);
        document.Pages[0].AddShape("2", document.Masters.Single(), 2, 2, 2, 2);
        document.Pages[0].AddShape("3", "Picture", 2, 2, 2, 2, unit: VisioMeasurementUnit.Inches);
        foreach (var candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            Assert.Equal(3, XDocument.Parse(candidate.Pages[0].ToSvg()).Descendants().Count(e => e.Attribute("data-officeimo-foreign-image") != null));
            // Isolate the imported instance for the resize/crop assertions.
            candidate.Pages[0].Shapes.RemoveAt(2); candidate.Pages[0].Shapes.RemoveAt(1);
            var shape = candidate.Pages[0].Shapes[0]; shape.Width = 1;
            Assert.True(VisioForeignImage.TryGetProjection(shape, candidate.Pages[0], 20, null, null, out var projection));
            Assert.Equal(20, projection.Width); Assert.Equal(0, projection.SourceCrop.Right);
            var raster = Render(candidate.Pages[0]);
            Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 25));
            Assert.Equal(OfficeColor.Green, raster.GetPixel(37, 25));
        }
    }

    [Fact]
    public void IndependentVdxBitmapRendersWithoutChangingOriginalPayload() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "angularjs-simple-scope.vdx");
        byte[] original = Convert.FromBase64String(XDocument.Load(path).Descendants(Ns + "ForeignData").Single().Value);
        Assert.Equal(0, BitConverter.ToInt32(original, 2));
        Assert.True(OfficeRasterImageDecoder.TryDecode(VisioForeignImage.PrepareBitmap(original), out var source));
        var document = VisioDocument.LoadLegacyXml(path).Value;
        foreach (var candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            var result = candidate.ExportImage(OfficeImageExportFormat.Svg, new VisioImageExportOptions { RenderText = false });
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "VISIO_BITMAP_SIZE_NORMALIZED");
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            var element = XDocument.Parse(Encoding.UTF8.GetString(result.Bytes)).Descendants().Single(e => e.Attribute("data-officeimo-foreign-image") != null);
            string data = element.Attributes().Single(attribute => attribute.Name.LocalName == "href").Value;
            Assert.True(OfficeRasterImageDecoder.TryDecode(Convert.FromBase64String(data.Substring(data.IndexOf(',') + 1)), out var actual));
            Assert.Equal(113, actual!.Width); Assert.Equal(155, actual.Height);
            for (int y = 0; y < actual.Height; y++)
                for (int x = 0; x < actual.Width; x++) Assert.Equal(source!.GetPixel(x, y), actual.GetPixel(x, y));
            var shape = candidate.Pages[0].Shapes.Single(s => s.Id == "41");
            shape.Width *= 1.5; shape.PinX -= 1;
            byte[] saved = candidate.ToLegacyXmlResult().Value;
            Assert.Equal(original, Convert.FromBase64String(XDocument.Load(new MemoryStream(saved)).Descendants(Ns + "ForeignData").Single().Value));
            Assert.Contains("data-officeimo-foreign-image", VisioDocument.LoadLegacyXml(new MemoryStream(saved)).Value.Pages[0].ToSvg());
        }
    }

    private sealed class RecordingCodec : IOfficeRasterImageCodec {
        public int Calls { get; private set; }
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            Calls++; image = null; return false;
        }
    }

    private static OfficeRasterImage Render(VisioPage page) {
        byte[] png = page.ToPng(new VisioPngSaveOptions { PixelsPerInch = 20, Supersampling = 1, RenderText = false, RenderStencilArtwork = false });
        Assert.True(OfficeRasterImageDecoder.TryDecode(png, out var raster)); return raster!;
    }
    private static VisioDocument Load(XDocument xml) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString()))).Value;
    private static MemoryStream Fixture() {
        var image = new OfficeRasterImage(16, 16, OfficeColor.White);
        for (int y = 0; y < 16; y++)
            for (int x = 0; x < 16; x++)
                image.SetPixel(x, y, y < 8 ? (x < 8 ? OfficeColor.Red : OfficeColor.Green) : (x < 8 ? OfficeColor.Blue : OfficeColor.White));
        var shape = new XElement(Ns + "Shape", new XAttribute("ID", "1"), new XAttribute("Type", "Foreign"),
            new XElement(Ns + "XForm", new XElement(Ns + "PinX", "2"), new XElement(Ns + "PinY", "2"), new XElement(Ns + "LocPinX", "1"), new XElement(Ns + "LocPinY", "1"), new XElement(Ns + "Width", "2"), new XElement(Ns + "Height", "2")),
            new XElement(Ns + "Foreign", new XElement(Ns + "ImgOffsetX", "0"), new XElement(Ns + "ImgOffsetY", "0"), new XElement(Ns + "ImgWidth", "2"), new XElement(Ns + "ImgHeight", "2")),
            new XElement(Ns + "ForeignData", new XAttribute("ForeignType", "Bitmap"), new XAttribute("CompressionType", "PNG"), Convert.ToBase64String(OfficePngWriter.Encode(image))));
        var xml = new XDocument(new XElement(Ns + "VisioDocument", new XElement(Ns + "Pages", new XElement(Ns + "Page", new XAttribute("ID", "0"), new XAttribute("NameU", "Image"),
            new XElement(Ns + "PageSheet", new XElement(Ns + "PageProps", new XElement(Ns + "PageWidth", "4"), new XElement(Ns + "PageHeight", "4"))), new XElement(Ns + "Shapes", shape)))));
        return new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString()));
    }
}
