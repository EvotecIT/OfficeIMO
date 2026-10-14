using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgOpacityTests {
    [Fact]
    public void PreservesIndependentFillAndStrokeOpacityAndProjectsThroughSharedDrawing() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        shape.FillColor = OdfColor.Parse("#FF0000"); shape.FillOpacity = .5; shape.StrokeOpacity = .25;
        var line = page.Shapes.AddLine(OdfLength.Centimeters(1), OdfLength.Centimeters(4), OdfLength.Centimeters(3), OdfLength.Centimeters(4));
        line.StrokeOpacity = .75;
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(.5, read.Pages[0].Shapes[0].FillOpacity); Assert.Equal(.25, read.Pages[0].Shapes[0].StrokeOpacity);
            var scene = read.Pages[0].ToDrawing().Value;
            var projected = scene.Elements.OfType<OfficeDrawingShape>().ToArray();
            Assert.Equal(.5, projected[0].Shape.FillOpacity); Assert.Equal(.25, projected[0].Shape.StrokeOpacity);
            Assert.Equal(.75, projected[1].Shape.StrokeOpacity); Assert.Null(projected[1].Shape.FillColor);
            Assert.Contains("fill-opacity=\"0.5\"", OfficeDrawingSvgExporter.ToSvg(scene));
            Assert.Contains("stroke-opacity=\"0.25\"", OfficeDrawingSvgExporter.ToSvg(scene));
            var raster = OfficeDrawingRasterRenderer.Render(scene);
            Assert.Equal(128, raster.GetPixel(40, 40).A);
        }
    }

    [Fact]
    public void InheritsDefaultsAndNamedOpacityAndRemovesOnlyTheLocalOverride() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        XElement styles = document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!;
        styles.Add(new XElement(OdfNamespaces.Style + "default-style", new XAttribute(OdfNamespaces.Style + "family", "graphic"),
            new XElement(OdfNamespaces.Style + "graphic-properties", new XAttribute(OdfNamespaces.Draw + "opacity", "60%"))));
        OdfStyle named = document.Styles.CreateNamed("Inherited", OdfStyleFamily.Graphic);
        named.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Svg + "stroke-opacity", "0.4");
        var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", named.Name); document.MarkPartDirty("content.xml");
        Assert.Equal(.6, shape.FillOpacity); Assert.Equal(.4, shape.StrokeOpacity);
        shape.FillOpacity = .2; shape.StrokeOpacity = .8;
        shape.FillOpacity = null; shape.StrokeOpacity = null;
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(.6, read.Pages[0].Shapes[0].FillOpacity); Assert.Equal(.4, read.Pages[0].Shapes[0].StrokeOpacity);
        }
    }

    [Fact]
    public void EditsSharedAutomaticStyleWithoutChangingTheOtherShape() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var first = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2)); first.FillOpacity = .5; first.StrokeOpacity = .25;
        var other = page.Shapes.AddRectangle(OdfRect.FromCentimeters(5, 1, 2, 2));
        other.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", first.Element.Attribute(OdfNamespaces.Draw + "style-name")!.Value);
        document.MarkPartDirty("content.xml");
        first.FillOpacity = .3;
        Assert.Equal(.3, first.FillOpacity); Assert.Equal(.5, other.FillOpacity);
        Assert.Equal(.25, first.StrokeOpacity); Assert.Equal(.25, other.StrokeOpacity);
    }

    [Theory]
    [InlineData(-.1)]
    [InlineData(1.1)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void RejectsInvalidEditsBeforeCreatingOrChangingStyles(double value) {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        string before = document.GetXml("content.xml").ToString();
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.FillOpacity = value);
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.StrokeOpacity = value);
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.ImageOpacity = value);
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    [Theory]
    [InlineData("opacity", "0.5")]
    [InlineData("opacity", "101%")]
    [InlineData("opacity", "1e1%")]
    [InlineData("stroke-opacity", "NaN")]
    public void InvalidNativeOpacityIsReportedAndStrictProjectionRejectsIt(string name, string value) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2), "Invalid");
        shape.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", (name == "opacity" ? OdfNamespaces.Draw : OdfNamespaces.Svg) + name, value);
        Assert.Empty(page.ToDrawing().Value.Elements);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Fact]
    public void PreservesOpacityGradientsAndReportsTheOpaqueProjectionFallback() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2), "Gradient");
        shape.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "opacity-name", "Fade");
        string before = document.GetXml("content.xml").ToString();
        Assert.Throws<NotSupportedException>(() => shape.FillOpacity = .5);
        Assert.Throws<NotSupportedException>(() => shape.FillOpacity);
        Assert.Equal(before, document.GetXml("content.xml").ToString());
        Assert.Single(page.ToDrawing().Value.Elements);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Fact]
    public void ProjectsImageOpacityWithoutChangingTheEmbeddedResource() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        var shape = page.Shapes.AddImage(png, "pixel.png", OdfRect.FromCentimeters(1, 1, 2, 2)); shape.ImageOpacity = .4; shape.FillOpacity = .8;
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(png, read.Pages[0].Shapes[0].GetImageBytes());
            Assert.Equal(.8, read.Pages[0].Shapes[0].FillOpacity);
            Assert.Equal(.4, read.Pages[0].Shapes[0].ImageOpacity);
            Assert.Equal(.4, Assert.Single(read.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingImage>()).Opacity);
        }
    }

    [Fact]
    public void ReportsGroupPaintOpacityWithoutDroppingItsVisibleChildren() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var group = page.Shapes.AddGroup("Transparent group"); group.FillOpacity = .5;
        group.Children.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        Assert.Single(page.ToDrawing().Value.Elements);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Fact]
    public void SharedOdfShapePropertiesAlsoPreservePresentationOpacity() {
        var presentation = OdpPresentation.Create();
        var shape = presentation.AddSlide("Opacity").AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        shape.FillOpacity = .2; shape.StrokeOpacity = .8;
        using var flat = new MemoryStream(); presentation.SaveFlatXml(flat); flat.Position = 0;
        foreach (var read in new[] { OdpPresentation.Load(new MemoryStream(presentation.ToBytes())), OdpPresentation.LoadFlatXml(flat) }) {
            Assert.Equal(.2, read.Slides[0].Shapes[0].FillOpacity);
            Assert.Equal(.8, read.Slides[0].Shapes[0].StrokeOpacity);
        }
    }

    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) };
    }
}
