using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgGeometryTests {
    [Theory]
    [InlineData("path")]
    [InlineData("polygon")]
    [InlineData("polyline")]
    public void NativeGeometryKeepsDeclaredViewBoxEdgesInsideTheProjectedCanvas(string kind) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var bounds = new OdfRect(OdfLength.Points(20), OdfLength.Points(20), OdfLength.Points(100), OdfLength.Points(100));
        var box = new OdfViewBox(0, 0, 11, 11);
        var points = new[] { new OfficePoint(0, 0), new OfficePoint(11, 0), new OfficePoint(11, 11), new OfficePoint(0, 11) };
        var shape = kind == "path" ? page.Shapes.AddPath(bounds, box, "M0 0L11 0 11 11 0 11Z") :
            kind == "polygon" ? page.Shapes.AddPolygon(bounds, box, points) : page.Shapes.AddPolyline(bounds, box, points);
        foreach (var read in new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes())) }) {
            var projected = Assert.Single(read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingShape>()).Shape;
            Assert.Equal(new OfficePoint(100, 100), projected.PathCommands[2].Point);
            Assert.Equal(100, projected.Width); Assert.Equal(100, projected.Height);
        }
        Assert.Equal(box.ToString(), shape.ViewBox.ToString());
    }

    [Fact]
    public void EditingTypedCommandsPreservesSmallCoordinatesAndControlPoints() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var bounds = OdfRect.FromCentimeters(1, 1, 20, 20); var box = new OdfViewBox(0, 0, 1, 1);
        var shape = page.Shapes.AddPath(bounds, box, "M0.0004 0.00000003Q0.123456789012345 0.987654321098765 0.999999999999999 0.00000009C0.2 0.3 0.4 0.5 0.6 0.7Z");
        var original = shape.PathCommands.ToArray();
        shape.SetPathCommands(original);
        var typed = page.Shapes.AddPath(bounds, box, original);
        var reopened = OdgDocument.Load(new MemoryStream(document.ToBytes()));
        foreach (var candidate in new[] { shape, typed, reopened.Pages[0].Shapes[0], reopened.Pages[0].Shapes[1] }) {
            Assert.Equal(original.Length, candidate.PathCommands.Count);
            for (int index = 0; index < original.Length; index++) {
                Assert.Equal(original[index].Point, candidate.PathCommands[index].Point);
                Assert.Equal(original[index].ControlPoint1, candidate.PathCommands[index].ControlPoint1);
                Assert.Equal(original[index].ControlPoint2, candidate.PathCommands[index].ControlPoint2);
            }
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ShapePaintUsesContentStyleScopeBeforeAndAfterEditing(bool drawing) {
        OdfDocument document = drawing ? OdgDocument.Create() : OdpPresentation.Create();
        var bounds = OdfRect.FromCentimeters(1, 1, 4, 2);
        OdfShape shape = drawing ? ((OdgDocument)document).AddPage().Shapes.AddRectangle(bounds)
            : ((OdpPresentation)document).AddSlide().AddRectangle(bounds);
        var style = document.Styles.CreateNamed("G", OdfStyleFamily.Graphic);
        XName properties = OdfNamespaces.Style + "graphic-properties";
        style.SetProperty(properties, OdfNamespaces.Svg + "fill-rule", "evenodd");
        style.SetProperty(properties, OdfNamespaces.Draw + "fill", "solid");
        style.SetProperty(properties, OdfNamespaces.Draw + "fill-color", "#FF0000");
        style.SetProperty(properties, OdfNamespaces.Draw + "stroke", "solid");
        style.SetProperty(properties, OdfNamespaces.Svg + "stroke-color", "#0000FF");
        style.SetProperty(properties, OdfNamespaces.Svg + "stroke-width", "2pt");
        var masterStyle = new XElement(style.Element);
        masterStyle.Element(properties)!.SetAttributeValue(OdfNamespaces.Svg + "fill-rule", "nonzero");
        masterStyle.Element(properties)!.SetAttributeValue(OdfNamespaces.Draw + "fill-color", "#00FF00");
        masterStyle.Element(properties)!.SetAttributeValue(OdfNamespaces.Svg + "stroke-color", "#00FF00");
        masterStyle.Element(properties)!.SetAttributeValue(OdfNamespaces.Svg + "stroke-width", "9pt");
        document.Package.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(masterStyle);
        document.Package.MarkXmlDirty("styles.xml");
        shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", "G"); document.Package.MarkXmlDirty("content.xml");
        Assert.Equal(OfficeFillRule.EvenOdd, shape.FillRule);
        Assert.Equal(OdfColor.Parse("#FF0000"), shape.FillColor);
        Assert.Equal(OdfColor.Parse("#0000FF"), shape.StrokeColor);
        Assert.Equal(2, shape.StrokeWidth!.Value.ToPoints());
        shape.StrokeWidth = OdfLength.Points(3);
        Assert.Equal(OfficeFillRule.EvenOdd, shape.FillRule);
        Assert.Equal(OdfColor.Parse("#FF0000"), shape.FillColor);
        Assert.Equal(OdfColor.Parse("#0000FF"), shape.StrokeColor);
        Assert.Equal(3, shape.StrokeWidth!.Value.ToPoints());
    }

    [Fact]
    public void IndependentLibreOfficePathsAndPolygonsKeepTheirLocalCanvasAndTransforms() {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-sheared-geometry.odg"));
        Assert.Equal(6, document.Pages.Count);
        foreach (OdgPage page in document.Pages) {
            OdgShape shape = Assert.Single(page.Shapes);
            var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            Assert.Single(result.Value.Elements.OfType<OfficeDrawingEffectGroup>());
            Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature == "page-style");
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "hairline-stroke");
            Assert.Contains("#FF00FF", OfficeDrawingSvgExporter.ToSvg(result.Value), StringComparison.OrdinalIgnoreCase);
            if (shape.ElementName == "rect") continue;
            Assert.Equal(5000, shape.ViewBox.Width); Assert.Equal(2000, shape.ViewBox.Height);
            if (shape.ElementName == "polygon") Assert.Equal(4, shape.Points.Count);
            else Assert.Contains(shape.PathCommands, command => command.Kind == OfficePathCommandKind.CubicBezierTo);
        }
        OdgShape curve = document.Pages[2].Shapes[0];
        string data = curve.PathData, transform = curve.Transform!;
        curve.Bounds = OdfRect.FromCentimeters(0, 0, 7, 3); curve.Name = "Edited curve";
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        var reopened = OdgDocument.LoadFlatXml(flat);
        Assert.Equal(data, reopened.Pages[2].Shapes[0].PathData);
        Assert.Equal(transform, reopened.Pages[2].Shapes[0].Transform);
        Assert.Equal(7, reopened.Pages[2].Shapes[0].Bounds.Width.ToCentimeters(), 3);
    }

    [Fact]
    public void PathEditingScalesTheDeclaredCanvasWithoutNormalizingAwayItsOrigin() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var bounds = new OdfRect(OdfLength.Points(10), OdfLength.Points(20), OdfLength.Points(80), OdfLength.Points(40));
        var shape = page.Shapes.AddPath(bounds, new OdfViewBox(-100, 50, 400, 200), "M-100 50 L300 50 L300 250 Z");
        shape.FillRule = OfficeFillRule.EvenOdd;
        var projected = Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal(10, projected.X); Assert.Equal(20, projected.Y);
        Assert.Equal(new OfficePoint(0, 0), projected.Shape.PathCommands[0].Point);
        Assert.Equal(new OfficePoint(80, 40), projected.Shape.PathCommands[2].Point);
        Assert.Equal(OfficeFillRule.EvenOdd, projected.Shape.FillRule);
        shape.SetPathCommands(new[] { OfficePathCommand.MoveTo(-100, 50), OfficePathCommand.QuadraticBezierTo(100, 150, 300, 250) });
        shape.Bounds = new OdfRect(bounds.X, bounds.Y, OdfLength.Points(160), bounds.Height);
        var read = OdgDocument.Load(new MemoryStream(document.ToBytes()));
        projected = Assert.Single(read.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal(OfficePathCommandKind.QuadraticBezierTo, projected.Shape.PathCommands[1].Kind);
        Assert.Equal(new OfficePoint(160, 40), projected.Shape.PathCommands[1].Point);
        Assert.Equal(new OfficePoint(80, 20), projected.Shape.PathCommands[1].ControlPoint1);
    }

    [Fact]
    public void PolygonAndPolylineEditingPreservesClosureAndDegeneratePageBounds() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var points = new[] { new OfficePoint(10, 20), new OfficePoint(110, 20), new OfficePoint(110, 120) };
        var polygon = page.Shapes.AddPolygon(OdfRect.FromCentimeters(1, 1, 4, 2), new OdfViewBox(10, 20, 100, 100), points);
        var line = page.Shapes.AddPolyline(OdfRect.FromCentimeters(1, 4, 4, 0), new OdfViewBox(10, 20, 100, 100), points);
        Assert.Null(line.FillColor);
        polygon.SetPoints(Enumerable.Reverse(points));
        var read = OdgDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(Enumerable.Reverse(points), read.Pages[0].Shapes[0].Points);
        var geometry = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingShape>().ToList();
        Assert.Equal(2, geometry.Count);
        Assert.Equal(OfficePathCommandKind.Close, geometry[0].Shape.PathCommands.Last().Kind);
        Assert.Equal(OfficePathCommandKind.LineTo, geometry[1].Shape.PathCommands.Last().Kind);
        Assert.All(geometry[1].Shape.PathCommands, command => Assert.Equal(0, command.Point.Y));
    }

    [Fact]
    public void RoundedRectangleRadiiFollowNativePrecedenceAndClampToTheCurrentBounds() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddRoundedRectangle(OdfRect.FromCentimeters(1, 1, 4, 2), OdfLength.Centimeters(5));
        var rounded = Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingShape>()).Shape;
        Assert.Equal(OfficeShapeKind.RoundedRectangle, rounded.Kind);
        Assert.Equal(OdfLength.Centimeters(1).ToPoints(), rounded.CornerRadius, 3);
        shape.CornerRadiusX = OdfLength.Centimeters(1); shape.CornerRadiusY = OdfLength.Centimeters(0.25);
        Assert.Null(shape.CornerRadius);
        var elliptical = Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingShape>()).Shape;
        Assert.Equal(OfficeShapeKind.Path, elliptical.Kind);
        Assert.Contains(elliptical.PathCommands, command => command.Kind == OfficePathCommandKind.CubicBezierTo);
        shape.CornerRadius = OdfLength.Points(5);
        Assert.Null(shape.CornerRadiusX); Assert.Null(shape.CornerRadiusY);
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.CornerRadius = OdfLength.Points(-1));
        Assert.Equal(5, shape.CornerRadius!.Value.ToPoints());
    }

    [Fact]
    public void InvalidOrExcessiveGeometryDoesNotReplaceExistingShapeData() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var box = new OdfViewBox(0, 0, 100, 100); var bounds = OdfRect.FromCentimeters(1, 1, 2, 2);
        var shape = page.Shapes.AddPath(bounds, box, "M0 0L100 100");
        Assert.Throws<InvalidDataException>(() => shape.PathData = "M0 0 Q1 2");
        Assert.ThrowsAny<ArgumentException>(() => shape.PathData = "M1e308 0l1e308 10");
        Assert.Equal("M0 0L100 100", shape.PathData);
        Assert.Throws<InvalidDataException>(() => shape.PathData = "M0 0" + string.Concat(Enumerable.Repeat("L1 1", 20000)));
        Assert.Throws<ArgumentException>(() => page.Shapes.AddPolyline(bounds, box, new[] { new OfficePoint(0, 0), new OfficePoint(0.5, 1) }));
        Assert.Throws<InvalidDataException>(() => page.Shapes.AddPolygon(bounds, box, Enumerable.Repeat(new OfficePoint(0, 0), 20001)));
        Assert.Throws<ArgumentException>(() => shape.ViewBox = default);
        Assert.Single(page.Shapes);
        shape.Element.SetAttributeValue(OdfNamespaces.Svg + "d", "invalid");
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }
}
