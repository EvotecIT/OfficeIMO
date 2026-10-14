using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgImageMirrorTests {
    [Theory]
    [InlineData(OdfImageMirror.None)]
    [InlineData(OdfImageMirror.Horizontal)]
    [InlineData(OdfImageMirror.Vertical)]
    [InlineData(OdfImageMirror.Horizontal | OdfImageMirror.Vertical)]
    [InlineData(OdfImageMirror.HorizontalOnOddPages)]
    [InlineData(OdfImageMirror.HorizontalOnEvenPages)]
    [InlineData(OdfImageMirror.HorizontalOnOddPages | OdfImageMirror.Vertical)]
    [InlineData(OdfImageMirror.HorizontalOnEvenPages | OdfImageMirror.Vertical)]
    public void LegalNativeModesSurviveBothContainersAndUnsupportedProjectionIsExplicit(OdfImageMirror mirror) {
        var (document, page, image, bytes) = Scene(); image.Mirror = mirror;
        foreach (OdgDocument read in RoundTrips(document)) {
            var shape = read.Pages[0].Shapes[0]; Assert.Equal(mirror, shape.Mirror); Assert.Equal(bytes, shape.GetImageBytes());
            var result = read.Pages[0].ToDrawing();
            bool supported = mirror is OdfImageMirror.None or OdfImageMirror.Horizontal;
            Assert.Equal(mirror == OdfImageMirror.Horizontal, Assert.Single(result.Value.Elements.OfType<OfficeDrawingImage>()).Projection.FlipHorizontal);
            if (supported) read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            else {
                Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":image-mirror") && m.Status == OdfConversionMappingStatus.Unsupported);
                Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            }
        }
    }

    [Fact]
    public void HorizontalMirrorReversesTheCroppedSourceWithoutChangingTheResource() {
        var (document, page, image, bytes) = Scene(); image.Crop = Insets(0, 30, 0, 0); image.Mirror = OdfImageMirror.Horizontal;
        foreach (OdgDocument read in RoundTrips(document)) {
            var drawing = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value;
            var projected = Assert.Single(drawing.Elements.OfType<OfficeDrawingImage>());
            Assert.True(projected.Projection.FlipHorizontal); Assert.Equal(.25, projected.Projection.SourceCrop.Right, 4);
            var raster = OfficeDrawingRasterRenderer.Render(drawing);
            Assert.Equal(OfficeColor.FromRgb(0, 0, 255), raster.GetPixel(45, 70));
            Assert.Equal(OfficeColor.FromRgb(255, 0, 0), raster.GetPixel(150, 70));
            Assert.Equal(bytes, read.Pages[0].Shapes[0].GetImageBytes());
            Assert.Contains("scale(-1 1)", OfficeDrawingSvgExporter.ToSvg(drawing));
        }
    }

    [Fact]
    public void NegativeCropMarginsReflectInsideTheOriginalFrameAndRetainOpacity() {
        var (_, page, image, _) = Scene(); image.Crop = Insets(0, 0, 0, -30); image.Mirror = OdfImageMirror.Horizontal; image.ImageOpacity = .5;
        var drawing = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value;
        var projected = Assert.Single(drawing.Elements.OfType<OfficeDrawingImage>());
        Assert.Equal(100, projected.Projection.RotationCenterX); Assert.Equal(70, projected.Projection.RotationCenterY);
        var raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(0, raster.GetPixel(150,70).A); Assert.Equal(128, raster.GetPixel(45,70).A);
    }

    [Fact]
    public void MirroringAnEmptyCropDoesNotCreatePixels() {
        var (_, page, image, _) = Scene(); image.Crop = Insets(0,-150,0,150); image.Mirror = OdfImageMirror.Horizontal;
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingImage>());
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":image-mirror") && m.Status == OdfConversionMappingStatus.Converted);
    }

    [Fact]
    public void InheritedMirrorCopiesOnWriteAndExplicitNoneDisablesTheParent() {
        var (document, page, image, bytes) = Scene();
        OdfStyle parent = document.Styles.CreateNamed("MirrorParent",OdfStyleFamily.Graphic);
        parent.SetProperty(OdfNamespaces.Style+"graphic-properties",OdfNamespaces.Style+"mirror","horizontal");
        image.Element.SetAttributeValue(OdfNamespaces.Draw+"style-name",parent.Name);
        var sibling = page.Shapes.AddImage(bytes,"sibling.png",image.Bounds); sibling.Element.SetAttributeValue(OdfNamespaces.Draw+"style-name",parent.Name);
        Assert.Equal(OdfImageMirror.Horizontal,image.Mirror); image.Mirror = OdfImageMirror.None;
        Assert.Equal(OdfImageMirror.Horizontal,sibling.Mirror); Assert.Equal(OdfImageMirror.None,image.Mirror);
        image.Mirror = null; Assert.Equal(OdfImageMirror.Horizontal,image.Mirror); Assert.Equal(bytes,image.GetImageBytes());
    }

    [Theory]
    [InlineData("diagonal")]
    [InlineData("none horizontal")]
    [InlineData("vertical vertical")]
    [InlineData("horizontal horizontal-on-odd")]
    public void MalformedMirrorIsPreservedWithAnUnmirroredCropFallback(string value) {
        var (document, page, image, bytes) = Scene(); image.Crop = Insets(0,30,0,0);
        image.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style+"graphic-properties",OdfNamespaces.Style+"mirror",value);
        string before = document.GetXml("content.xml").ToString();
        Assert.Throws<InvalidDataException>(() => image.Mirror);
        var result = page.ToDrawing(); var projected = Assert.Single(result.Value.Elements.OfType<OfficeDrawingImage>());
        Assert.False(projected.Projection.FlipHorizontal); Assert.Equal(.25,projected.Projection.SourceCrop.Right,4);
        Assert.Contains(result.Report.Mappings,m=>m.Feature.EndsWith(":image-mirror") && m.Status==OdfConversionMappingStatus.Unsupported);
        Assert.Equal(bytes,image.GetImageBytes()); Assert.Equal(before,document.GetXml("content.xml").ToString());
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Fact]
    public void InvalidMirrorEditsFailBeforeStyleMutationAndNonImagesRejectAccess() {
        var (document,page,image,_) = Scene(); string before = document.GetXml("content.xml").ToString();
        Assert.Throws<ArgumentOutOfRangeException>(()=>image.Mirror=OdfImageMirror.Horizontal|OdfImageMirror.HorizontalOnOddPages);
        Assert.Throws<ArgumentOutOfRangeException>(()=>image.Mirror=(OdfImageMirror)16); Assert.Equal(before,document.GetXml("content.xml").ToString());
        var rectangle=page.Shapes.AddRectangle(image.Bounds);
        Assert.Throws<InvalidOperationException>(()=>rectangle.Mirror); Assert.Throws<InvalidOperationException>(()=>rectangle.Mirror=null);
    }

    [Fact]
    public void DefaultGraphicMirrorAndReversedTokenOrderAreResolved() {
        var (document,_,image,_) = Scene();
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office+"styles")!.Add(new XElement(OdfNamespaces.Style+"default-style",
            new XAttribute(OdfNamespaces.Style+"family","graphic"),new XElement(OdfNamespaces.Style+"graphic-properties",new XAttribute(OdfNamespaces.Style+"mirror","horizontal"))));
        Assert.Equal(OdfImageMirror.Horizontal,image.Mirror);
        image.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style+"graphic-properties",OdfNamespaces.Style+"mirror","vertical\t horizontal-on-even");
        Assert.Equal(OdfImageMirror.Vertical|OdfImageMirror.HorizontalOnEvenPages,image.Mirror);
    }

    private static (OdgDocument,OdgPage,OdgShape,byte[]) Scene() {
        var document=OdgDocument.Create();var page=document.AddPage();var raster=new OfficeRasterImage(160,80);var canvas=new OfficeRasterCanvas(raster);
        OfficeColor[] colors={OfficeColor.FromRgb(255,0,0),OfficeColor.FromRgb(0,255,0),OfficeColor.FromRgb(0,0,255),OfficeColor.FromRgb(255,255,0)};
        for(int i=0;i<colors.Length;i++)canvas.FillRectangle(i*40,0,40,80,colors[i]);
        byte[] bytes=OfficePngWriter.Encode(raster,new OfficePngEncodeOptions{DpiX=96,DpiY=96});
        return(document,page,page.Shapes.AddImage(bytes,"mirror.png",new OdfRect(P(40),P(40),P(120),P(60)),"Mirror"),bytes);
    }
    private static OdfLength P(double n)=>OdfLength.Points(n);
    private static OdfInsets Insets(double top,double right,double bottom,double left)=>new OdfInsets(P(top),P(right),P(bottom),P(left));
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat=new MemoryStream();document.SaveFlatXml(flat);flat.Position=0;
        return new[]{OdgDocument.Load(new MemoryStream(document.ToBytes())),OdgDocument.LoadFlatXml(flat)};
    }
}
