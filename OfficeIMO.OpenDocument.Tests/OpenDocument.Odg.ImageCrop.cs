using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgImageCropTests {
    [Fact]
    public void CropUsesIntrinsicImageSizeAndSurvivesFrameResizeAndBothContainers() {
        var (document, page, shape, bytes) = Scene();
        shape.Crop = Insets(7.5, 7.5, 7.5, 15);
        foreach (OdgDocument read in RoundTrips(document)) {
            OdgShape image = read.Pages[0].Shapes[0];
            Assert.Equal(bytes, image.GetImageBytes()); Assert.Equal(shape.Crop, image.Crop);
            AssertProjection(read.Pages[0], .125, .125, .0625, .125);
            image.Bounds = new OdfRect(P(40), P(40), P(160), P(60));
            AssertProjection(read.Pages[0], .125, .125, .0625, .125);
            var raster = OfficeDrawingRasterRenderer.Render(read.Pages[0].ToDrawing().Value);
            Assert.Equal(OfficeColor.FromRgb(0, 255, 0), raster.GetPixel(45, 70));
            Assert.Equal(OfficeColor.FromRgb(0, 0, 255), raster.GetPixel(70, 70));
            Assert.Contains("clipPath", OfficeDrawingSvgExporter.ToSvg(read.Pages[0].ToDrawing().Value));
        }
    }

    [Fact]
    public void NativeCropReadsStylesAndCopiesOnWriteWithoutChangingSiblingOrResource() {
        var (document, page, shape, bytes) = Scene();
        OdfStyle parent = document.Styles.CreateNamed("ParentCrop", OdfStyleFamily.Graphic);
        parent.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "clip", "rect(auto, 7.5pt, auto, 15pt)");
        shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", parent.Name);
        OdgShape sibling = page.Shapes.AddImage(bytes, "sibling.png", shape.Bounds);
        sibling.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", parent.Name);
        Assert.Equal(new OdfInsets(default, P(7.5), default, P(15)), shape.Crop);
        shape.Crop = Insets(0, 0, 0, 30);
        Assert.Equal(new OdfInsets(default, P(7.5), default, P(15)), sibling.Crop);
        Assert.Equal(bytes, shape.GetImageBytes());
        shape.Crop = null; Assert.Equal(sibling.Crop, shape.Crop);
        shape.Crop = default(OdfInsets); Assert.Equal(default(OdfInsets), shape.Crop);
        AssertProjection(page, 0, 0, 0, 0);
    }

    [Theory]
    [InlineData("auto")]
    [InlineData("rect(auto auto auto auto)")]
    [InlineData("rect(0cm, 0in, 0mm, 0pt)")]
    public void ExplicitAutoOrZeroDisablesInheritedCrop(string lexical) {
        var (document, page, shape, _) = Scene();
        XElement common = document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!;
        common.Add(new XElement(OdfNamespaces.Style + "default-style", new XAttribute(OdfNamespaces.Style + "family", "graphic"),
            new XElement(OdfNamespaces.Style + "graphic-properties", new XAttribute(OdfNamespaces.Fo + "clip", "rect(0pt 0pt 0pt 15pt)"))));
        Assert.Equal(15, shape.Crop!.Value.Left.ToPoints());
        shape.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "clip", lexical);
        AssertProjection(page, 0, 0, 0, 0);
    }

    [Fact]
    public void NegativeCropKeepsEmptyMarginsInsteadOfStretchingPixelsAcrossTheFrame() {
        var (_, page, shape, _) = Scene(); shape.Crop = Insets(-15, 0, 0, -30); shape.ImageOpacity = .5;
        OfficeDrawingImage image = Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingImage>());
        // PNG's integer pixels-per-meter metadata introduces sub-point quantization.
        Assert.InRange(image.Projection.X, 63.99, 64.01); Assert.InRange(image.Projection.Y, 51.99, 52.01);
        Assert.InRange(image.Projection.Width, 95.99, 96.01); Assert.InRange(image.Projection.Height, 47.99, 48.01);
        Assert.False(image.Projection.HasCrop);
        var raster = OfficeDrawingRasterRenderer.Render(page.ToDrawing().Value);
        Assert.Equal(0, raster.GetPixel(50, 60).A); Assert.Equal(128, raster.GetPixel(70, 65).A);
    }

    [Fact]
    public void EmptyIntersectionDoesNotInventImagePixels() {
        var (_, page, shape, _) = Scene(); shape.Crop = Insets(0, -150, 0, 150);
        var projected = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.Empty(projected.Value.Elements.OfType<OfficeDrawingImage>());
        Assert.Contains(projected.Report.Mappings, m => m.Feature.EndsWith(":image-crop") && m.Status == OdfConversionMappingStatus.Converted);
    }

    [Theory]
    [InlineData("rect(0cm 0cm 0cm)")]
    [InlineData("rect(0cm 0cm 0cm 10%)")]
    [InlineData("rect(0cm 0cm 0cm NaNpt)")]
    public void InvalidImportedCropIsPreservedAndReportedWithAnUncroppedFallback(string lexical) {
        var (_, page, shape, bytes) = Scene();
        shape.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "clip", lexical);
        Assert.Throws<InvalidDataException>(() => shape.Crop);
        var projected = page.ToDrawing(); Assert.Single(projected.Value.Elements.OfType<OfficeDrawingImage>());
        Assert.Contains(projected.Report.Mappings, m => m.Feature.EndsWith(":image-crop") && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Equal(bytes, shape.GetImageBytes());
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Fact]
    public void InvalidEditsFailBeforeStyleMutationAndNonImagesRejectCrop() {
        var (document, page, shape, _) = Scene(); string before = document.GetXml("content.xml").ToString();
        Assert.Throws<InvalidDataException>(() => shape.Crop = new OdfInsets(P(0), P(0), P(0), OdfLength.Parse("1em")));
        Assert.Equal(before, document.GetXml("content.xml").ToString());
        var rectangle = page.Shapes.AddRectangle(shape.Bounds);
        Assert.Throws<InvalidOperationException>(() => rectangle.Crop); Assert.Throws<InvalidOperationException>(() => rectangle.Crop = null);
    }

    [Fact]
    public void CollapsedCropAndInvalidMirroringRemainExplicitProjectionLosses() {
        var (_, page, shape, _) = Scene(); shape.Crop = Insets(0, 0, 0, 120);
        shape.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Style + "mirror", "diagonal");
        var result = page.ToDrawing(); Assert.Single(result.Value.Elements.OfType<OfficeDrawingImage>());
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":image-crop") && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":image-mirror") && m.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void PresentationCropSharesTheInheritedNativeCodec() {
        var presentation = OdpPresentation.Create(); OdpImage image = presentation.AddSlide().AddImage(Grid(), "grid.png", new OdfRect(P(0), P(0), P(120), P(60)));
        OdfStyle parent = presentation.Styles.CreateNamed("ParentCrop", OdfStyleFamily.Graphic);
        parent.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "clip", "rect(0pt 0pt 0pt 15pt)");
        image.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", parent.Name);
        Assert.Equal(15, image.Crop!.Value.Left.ToPoints()); image.Crop = Insets(0, 0, 0, 30);
        Assert.Equal(30, image.Crop!.Value.Left.ToPoints()); image.Crop = null; Assert.Equal(15, image.Crop!.Value.Left.ToPoints());
    }

    [Fact]
    public void SourceFractionsUseEmbeddedHorizontalAndVerticalDensity() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var image = page.Shapes.AddImage(Grid(192, 96), "dense.png", new OdfRect(P(40), P(40), P(120), P(60)));
        image.Crop = Insets(7.5, 7.5, 7.5, 15);
        foreach (OdgDocument read in RoundTrips(document)) AssertProjection(read.Pages[0], .25, .125, .125, .125);
    }

    [Fact]
    public void NativeDecimalOffsetsKeepTheProjectionInsideItsFrame() {
        var (_, page, shape, _) = Scene();
        shape.Crop = OdfInsets.FromCentimeters(.529, 0, .265, 0);
        var image = Assert.Single(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingImage>());
        Assert.Equal(60, image.Projection.Height); Assert.Equal(120, image.Projection.Width);
    }

    private static void AssertProjection(OdgPage page, double left, double top, double right, double bottom) {
        OfficeDrawingImage image = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingImage>().First();
        Assert.Equal(left, image.Projection.SourceCrop.Left, 4); Assert.Equal(top, image.Projection.SourceCrop.Top, 4);
        Assert.Equal(right, image.Projection.SourceCrop.Right, 4); Assert.Equal(bottom, image.Projection.SourceCrop.Bottom, 4);
    }
    private static (OdgDocument, OdgPage, OdgShape, byte[]) Scene() {
        OdgDocument document = OdgDocument.Create(); OdgPage page = document.AddPage(); byte[] bytes = Grid();
        return (document, page, page.Shapes.AddImage(bytes, "grid.png", new OdfRect(P(40), P(40), P(120), P(60)), "Crop"), bytes);
    }
    private static byte[] Grid(double dpiX = 96, double dpiY = 96) {
        var image = new OfficeRasterImage(160, 80); var canvas = new OfficeRasterCanvas(image);
        OfficeColor[] colors = { OfficeColor.FromRgb(255, 0, 0), OfficeColor.FromRgb(0, 255, 0), OfficeColor.FromRgb(0, 0, 255), OfficeColor.FromRgb(255, 255, 0), OfficeColor.FromRgb(0, 255, 255), OfficeColor.FromRgb(255, 0, 255), OfficeColor.FromRgb(128, 64, 32), OfficeColor.FromRgb(64, 128, 32) };
        for (int x = 0; x < colors.Length; x++) canvas.FillRectangle(x * 20, 0, 20, 80, colors[x]);
        return OfficePngWriter.Encode(image, new OfficePngEncodeOptions { DpiX = dpiX, DpiY = dpiY });
    }
    private static OdfLength P(double value) => OdfLength.Points(value);
    private static OdfInsets Insets(double top, double right, double bottom, double left) => new OdfInsets(P(top), P(right), P(bottom), P(left));
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) };
    }
}
