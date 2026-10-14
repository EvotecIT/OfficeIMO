using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgDiagramTests {
    [Fact]
    public void ReadsAndEditsIndependentLibreOfficeLayerState() {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-layers.odg"));
        Assert.Equal(7, document.Layers.Count);
        Assert.Equal(OdgLayerDisplay.Screen, document.Layers.Find("V--")!.Display);
        Assert.True(document.Layers.Find("V-L")!.IsProtected);
        Assert.Empty(document.Pages[0].ToDrawing(forPrint: true).Value.Elements);
        document.Layers.Find("V--")!.Display = OdgLayerDisplay.Printer;
        document.Layers.Find("V-L")!.IsProtected = false;
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        var reopened = OdgDocument.LoadFlatXml(flat);
        Assert.Equal(OdgLayerDisplay.Printer, reopened.Layers.Find("V--")!.Display);
        Assert.False(reopened.Layers.Find("V-L")!.IsProtected);
        // The legacy producer reads its saved-view masks; confirm the edit survives without explicit attributes too.
        foreach (var layer in reopened.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "layer")) {
            layer.Attribute(OdfNamespaces.Draw + "display")?.Remove(); layer.Attribute(OdfNamespaces.Draw + "protected")?.Remove();
        }
        Assert.Equal(OdgLayerDisplay.Printer, reopened.Layers.Find("V--")!.Display);
        Assert.False(reopened.Layers.Find("V-L")!.IsProtected);
    }

    [Fact]
    public void ConnectorsFollowMovedResizedAndRenamedShapesAndDetachOnRemoval() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var first = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        var second = page.Shapes.AddRectangle(OdfRect.FromCentimeters(10, 2, 4, 2));
        var connector = page.Shapes.AddConnector(first.AddGluePoint(OdgGluePointAlignment.Right), second.AddGluePoint(OdgGluePointAlignment.Left));
        Assert.Equal(3, connector.X1.ToCentimeters(), 3);
        first.Bounds = OdfRect.FromCentimeters(2, 3, 4, 4);
        first.Element.SetAttributeValue(OdfNamespaces.Draw + "id", first.XmlId);
        first.Element.Attribute(XNamespace.Xml + "id")!.Remove(); // Older producer ID spelling.
        first.XmlId = "renamed";
        Assert.Equal("renamed", connector.StartShapeId);
        Assert.Equal(6, connector.X1.ToCentimeters(), 3); Assert.Equal(5, connector.Y1.ToCentimeters(), 3);
        Assert.Throws<InvalidOperationException>(() => connector.X1 = OdfLength.Points(50));
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        var reopened = OdgDocument.LoadFlatXml(flat); var read = reopened.Pages[0].Shapes[2];
        Assert.Equal(6, read.X1.ToCentimeters(), 3);
        reopened.Pages[0].Shapes.RemoveAt(0);
        Assert.Null(read.StartShapeId); Assert.Equal(6, read.X1.ToCentimeters(), 3);
        Assert.NotNull(read.EndShapeId);
        read.AttachEnd(null); read.X2 = OdfLength.Centimeters(15);
        Assert.Equal(15, read.X2.ToCentimeters(), 3);
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void RejectsCrossPageAttachmentsWithoutLeavingAConnector() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var other = document.AddPage();
        var first = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2)).AddGluePoint(OdgGluePointAlignment.Right);
        var second = other.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2)).AddGluePoint(OdgGluePointAlignment.Left);
        Assert.Throws<ArgumentException>(() => page.Shapes.AddConnector(first, second));
        Assert.Single(page.Shapes);
        Assert.Null(page.Shapes[0].XmlId);
    }

    [Fact]
    public void ProjectsLayerInheritanceAndPrintIntentWithoutTreatingDefinitionsAsShapes() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        document.Layers.Add("Screen", OdgLayerDisplay.Screen); document.Layers.Add("Print", OdgLayerDisplay.Printer);
        document.Layers.Add("Hidden", OdgLayerDisplay.None);
        foreach (string name in new[] { "Screen", "Print", "Hidden" }) {
            var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2), name); shape.Layer = name; shape.Text = name;
        }
        Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Contains("Screen", OfficeDrawingSvgExporter.ToSvg(page.ToDrawing().Value));
        Assert.DoesNotContain(">Hidden<", OfficeDrawingSvgExporter.ToSvg(page.ToDrawing().Value));
        Assert.Contains("Print", OfficeDrawingSvgExporter.ToSvg(page.ToDrawing(forPrint: true).Value));
        page.Layers.Add("Hidden", OdgLayerDisplay.Always);
        Assert.Equal(3, page.Shapes.Count);
        Assert.Contains("Hidden", OfficeDrawingSvgExporter.ToSvg(page.ToDrawing().Value));
        var read = OdgDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Single(read.Pages[0].EffectiveLayers); Assert.Equal(3, read.Layers.Count);
    }

    [Fact]
    public void AppliesNativeRadianTransformsToShapesAndAttachedConnectors() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var rotated = page.Shapes.AddRectangle(new OdfRect(OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(20), OdfLength.Points(10)));
        rotated.Transform = "rotate(-1.5707963267948966) translate(100pt 50pt)";
        var target = page.Shapes.AddEllipse(new OdfRect(OdfLength.Points(150), OdfLength.Points(60), OdfLength.Points(20), OdfLength.Points(20)));
        var connector = page.Shapes.AddConnector(rotated.AddGluePoint(OdgGluePointAlignment.Right), target.AddGluePoint(OdgGluePointAlignment.Left));
        Assert.Equal(95, connector.X1.ToPoints(), 3); Assert.Equal(70, connector.Y1.ToPoints(), 3);
        var result = page.ToDrawing();
        Assert.False(result.Report.HasSkippedOrUnsupported);
        var transformed = Assert.Single(result.Value.Elements.OfType<OfficeDrawingEffectGroup>());
        var painted = Assert.Single(transformed.Drawing.Elements.OfType<OfficeDrawingShape>());
        var origin = transformed.Transform.TransformPoint(new OfficePoint(painted.X, painted.Y));
        Assert.Equal(100, origin.X, 3); Assert.Equal(50, origin.Y, 3);
        Assert.Contains("matrix", OfficeDrawingSvgExporter.ToSvg(result.Value));
        rotated.Transform = "unsupported(1)";
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Fact]
    public void PreservesOdfChildOrderWhenAddingGluePointsAndPageLayers() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        page.Element.AddFirst(new XElement(OdfNamespaces.Svg + "title", "Title"), new XElement(OdfNamespaces.Svg + "desc", "Description"));
        page.Layers.Add("Visible");
        Assert.Equal(new[] { "title", "desc", "layer-set" }, page.Element.Elements().Select(element => element.Name.LocalName));
        foreach (var shape in new[] { page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2)), page.Shapes.AddEllipse(OdfRect.FromCentimeters(5, 1, 2, 2)) }) {
            shape.Text = "Label"; shape.AddGluePoint(OdgGluePointAlignment.Right);
            Assert.Equal(new[] { "glue-point", "p" }, shape.ToXml().Elements().Select(element => element.Name.LocalName));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DetachesNativeAutomaticConnectorUsingSavedEndpoint(bool removeShape) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var a = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2)); var b = page.Shapes.AddEllipse(OdfRect.FromCentimeters(5, 1, 2, 2));
        var connector = page.Shapes.AddConnector(a.AddGluePoint(), b.AddGluePoint());
        connector.Element.Attribute(OdfNamespaces.Draw + "start-glue-point")!.Remove();
        connector.Element.SetAttributeValue(OdfNamespaces.Svg + "x1", "3cm"); connector.Element.SetAttributeValue(OdfNamespaces.Svg + "y1", "2cm");
        if (removeShape) page.Shapes.RemoveAt(0); else connector.AttachStart(null);
        Assert.Null(connector.StartShapeId); Assert.Equal(3, connector.X1.ToCentimeters(), 3); Assert.Equal(2, connector.Y1.ToCentimeters(), 3);
    }

    [Fact]
    public void ProjectsLinesOnTheBottomAndRightPageBoundaries() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        page.Shapes.AddLine(OdfLength.Points(0), page.Height, page.Width, page.Height);
        page.Shapes.AddLine(page.Width, OdfLength.Points(0), page.Width, page.Height);
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.Equal(2, result.Value.Elements.Count);
    }

    [Fact]
    public void PreservesAndReportsUnsupportedConnectorRouting() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var a = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2)); var b = page.Shapes.AddEllipse(OdfRect.FromCentimeters(5, 1, 2, 2));
        var connector = page.Shapes.AddConnector(a.AddGluePoint(OdgGluePointAlignment.Right), b.AddGluePoint(OdgGluePointAlignment.Left));
        connector.ConnectorKind = OdgConnectorKind.Curve;
        var reopened = OdgDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(OdgConnectorKind.Curve, reopened.Pages[0].Shapes[2].ConnectorKind);
        Assert.Throws<OdfConversionLossException>(() => reopened.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }
}
