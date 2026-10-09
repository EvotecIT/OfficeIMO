using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgConnectorRouteTests {
    [Fact]
    public void ReadsIndependentNativeRoutesAndLegacyRelativeGlueWithoutChangingXml() {
        var document = LoadFixture("libreoffice-routed-glue.odg");
        var connectors = document.Pages[0].Shapes.Where(shape => shape.IsConnector).ToArray();
        Assert.Equal(2, connectors.Length);
        foreach (var connector in connectors) {
            string before = connector.ToXml().ToString();
            var route = connector.ConnectorRouteCommands;
            Assert.Equal(4, route.Count);
            Assert.Equal(OfficePathCommandKind.MoveTo, route[0].Kind);
            Assert.All(route.Skip(1), command => Assert.Equal(OfficePathCommandKind.LineTo, command.Kind));
            AssertClose(new OfficePoint(connector.X1.ToPoints(), connector.Y1.ToPoints()), route[0].Point);
            AssertClose(new OfficePoint(connector.X2.ToPoints(), connector.Y2.ToPoints()), route[route.Count - 1].Point);
            Assert.Equal(before, connector.ToXml().ToString());
        }
        Assert.Equal(11.699, connectors[0].X2.ToCentimeters(), 2);
        // This producer's embedded SVG lies outside the flat-XML image profile; preserve it in package output.
        var read = OdgDocument.Load(new MemoryStream(document.ToBytes())).Pages[0].Shapes.Where(shape => shape.IsConnector).ToArray();
        Assert.Equal(connectors[0].ConnectorRouteCommands, read[0].ConnectorRouteCommands);
    }

    [Fact]
    public void RetainsIndependentAutomaticEdgeAndSkewedRoute() {
        // The original ZIP has an extra mimetype field. Repack without changing any XML/resource entry bytes.
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-network-connectors.odg"));
        var document = OdgDocument.Load(new MemoryStream(OdfTestPackageRewriter.Rewrite(source)));
        var connector = document.Pages[0].Shapes.Single(shape => shape.XmlId == "id15");
        Assert.Null(connector.Element.Attribute(OdfNamespaces.Draw + "end-glue-point"));
        Assert.Equal(18.219, connector.X2.ToCentimeters(), 3);
        Assert.Equal(9.63, connector.Y2.ToCentimeters(), 3);
        var route = connector.ConnectorRouteCommands;
        Assert.Equal(4, route.Count);
        Assert.Equal(18.7, route[1].Point.Y * 2.54 / 72, 3);
        AssertClose(new OfficePoint(connector.X2.ToPoints(), connector.Y2.ToPoints()), route[3].Point);
    }

    [Theory]
    [InlineData(OdgConnectorKind.Standard)]
    [InlineData(OdgConnectorKind.Lines)]
    [InlineData(OdgConnectorKind.Curve)]
    public void AuthorsAndProjectsOpenRoutesAcrossFlatAndPackageOutput(OdgConnectorKind kind) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(20, 30), new OfficePoint(140, 90), kind);
        var commands = kind == OdgConnectorKind.Curve
            ? new[] { OfficePathCommand.MoveTo(20, 30), OfficePathCommand.CubicBezierTo(60, 10, 100, 120, 140, 90) }
            : new[] { OfficePathCommand.MoveTo(20, 30), OfficePathCommand.LineTo(70, 30), OfficePathCommand.LineTo(70, 90), OfficePathCommand.LineTo(140, 90) };
        connector.SetConnectorRoute(commands); connector.FillColor = OdfColor.Parse("FF0000");
        foreach (var read in new[] { OdgDocument.Load(new MemoryStream(document.ToBytes())), FlatRoundTrip(document) }) {
            var shape = read.Pages[0].Shapes[0];
            Assert.Equal(commands.Length, shape.ConnectorRouteCommands.Count);
            for (int i = 0; i < commands.Length; i++) {
                Assert.Equal(commands[i].Kind, shape.ConnectorRouteCommands[i].Kind);
                AssertClose(commands[i].Point, shape.ConnectorRouteCommands[i].Point);
            }
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var projected = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
            Assert.Null(projected.Shape.FillColor);
            Assert.Contains(kind == OdgConnectorKind.Curve ? "C" : "L", OfficeDrawingSvgExporter.ToSvg(result.Value));
        }
    }

    [Fact]
    public void AutomaticAttachmentsFollowTransformedShapesAndDetachAtTheirResolvedPosition() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var group = page.Shapes.AddGroup();
        var start = group.Children.AddRectangle(Rect(0, 0, 20, 10));
        group.TransformChildren("scale(2 1) translate(50pt 20pt)");
        var end = page.Shapes.AddRectangle(Rect(150, 60, 20, 20));
        var connector = page.Shapes.AddConnector(new OfficePoint(0, 0), new OfficePoint(150, 70));
        connector.AttachStartToShape(start); connector.AttachEndToShape(end);
        AssertClose(new OfficePoint(90, 25), new OfficePoint(connector.X1.ToPoints(), connector.Y1.ToPoints()));
        AssertClose(new OfficePoint(150, 70), new OfficePoint(connector.X2.ToPoints(), connector.Y2.ToPoints()));
        start.Bounds = Rect(10, 20, 40, 20);
        AssertClose(new OfficePoint(150, 50), new OfficePoint(connector.X1.ToPoints(), connector.Y1.ToPoints()));
        connector.AttachStartToShape(null);
        Assert.Null(connector.StartShapeId); Assert.Equal(150, connector.X1.ToPoints(), 3);
        Assert.False(page.ToDrawing().Report.HasSkippedOrUnsupported);
    }

    [Fact]
    public void RejectsStaleAttachedRouteUntilCallerReplacesIt() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(10, 10, 20, 20)); var end = page.Shapes.AddRectangle(Rect(100, 50, 20, 20));
        var connector = page.Shapes.AddConnector(start.AddGluePoint(OdgGluePointAlignment.Right), end.AddGluePoint(OdgGluePointAlignment.Left));
        connector.ConnectorKind = OdgConnectorKind.Standard;
        connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(30, 20), OfficePathCommand.LineTo(60, 20), OfficePathCommand.LineTo(60, 60), OfficePathCommand.LineTo(100, 60) });
        start.Bounds = Rect(10, 30, 20, 20);
        Assert.Throws<NotSupportedException>(() => connector.ConnectorRouteCommands);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(30, 40), OfficePathCommand.LineTo(60, 40), OfficePathCommand.LineTo(60, 60), OfficePathCommand.LineTo(100, 60) });
        Assert.False(page.ToDrawing().Report.HasSkippedOrUnsupported);
    }

    [Fact]
    public void SupportsDeclaredViewBoxCachesAndClearsThemAfterEndpointOrKindEdits() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(30, 40), new OfficePoint(130, 90), OdgConnectorKind.Curve);
        connector.Element.SetAttributeValue(OdfNamespaces.Svg + "viewBox", "-10 -20 100 50");
        connector.Element.SetAttributeValue(OdfNamespaces.Svg + "d", "M-10 -20C20 -20 40 30 90 30");
        var route = connector.ConnectorRouteCommands;
        AssertClose(new OfficePoint(30, 40), route[0].Point); AssertClose(new OfficePoint(130, 90), route[1].Point);
        connector.X2 = OdfLength.Points(150);
        Assert.Null(connector.Element.Attribute(OdfNamespaces.Svg + "d"));
        Assert.Throws<NotSupportedException>(() => connector.ConnectorRouteCommands);
        connector.ConnectorKind = OdgConnectorKind.Line;
        AssertClose(new OfficePoint(150, 90), connector.ConnectorRouteCommands[1].Point);
    }

    [Fact]
    public void InvalidRouteEditsAreAtomicAndBounded() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(10, 20), new OfficePoint(100, 20));
        string before = connector.ToXml().ToString();
        Assert.Throws<NotSupportedException>(() => connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(10, 10), OfficePathCommand.Close() }));
        Assert.Equal(before, connector.ToXml().ToString());
        Assert.Throws<InvalidDataException>(() => connector.SetConnectorRoute(Enumerable.Repeat(OfficePathCommand.LineTo(1, 1), 20001).Prepend(OfficePathCommand.MoveTo(0, 0))));
        Assert.Equal(before, connector.ToXml().ToString());
        var a = page.Shapes.AddRectangle(Rect(1, 1, 20, 20)); connector.AttachStart(a.AddGluePoint()); before = connector.ToXml().ToString();
        Assert.Throws<ArgumentException>(() => connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(100, 20) }));
        Assert.Equal(before, connector.ToXml().ToString());
    }

    [Theory]
    [InlineData(false, OdgConnectorKind.Line)]
    [InlineData(true, OdgConnectorKind.Line)]
    [InlineData(false, OdgConnectorKind.Curve)]
    [InlineData(true, OdgConnectorKind.Curve)]
    public void AttachedEndpointsRemainValidAfterNativeGridRounding(bool automatic, OdgConnectorKind kind) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(10, 10, 20, 20));
        var end = page.Shapes.AddRectangle(Rect(100, 10, 20, 20));
        var connector = page.Shapes.AddConnector(new OfficePoint(30, 20), new OfficePoint(100, 20), kind);
        if (automatic) { connector.AttachStartToShape(start); connector.AttachEndToShape(end); }
        else { connector.AttachStart(start.AddGluePoint(OdgGluePointAlignment.Right)); connector.AttachEnd(end.AddGluePoint(OdgGluePointAlignment.Left)); }
        var last = kind == OdgConnectorKind.Curve
            ? OfficePathCommand.CubicBezierTo(50, 0, 80, 40, 99.905, 20)
            : OfficePathCommand.LineTo(99.905, 20);
        connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(30.095, 20), last });
        foreach (var read in new[] { document, FlatRoundTrip(document), OdgDocument.Load(new MemoryStream(document.ToBytes())) }) {
            var saved = read.Pages[0].Shapes.Single(shape => shape.IsConnector);
            var route = saved.ConnectorRouteCommands;
            AssertClose(new OfficePoint(30, 20), route[0].Point);
            AssertClose(new OfficePoint(100, 20), route[1].Point);
            Assert.Equal(last.Kind, route[1].Kind);
            if (kind == OdgConnectorKind.Curve) {
                AssertClose(last.ControlPoint1, route[1].ControlPoint1);
                AssertClose(last.ControlPoint2, route[1].ControlPoint2);
            }
            Assert.False(read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Report.HasSkippedOrUnsupported);
        }
    }

    [Fact]
    public void GroupTransformsBakeFreeRoutesAndPreserveNestedShapeTransforms() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var group = page.Shapes.AddGroup();
        var nested = group.Children.AddGroup(); var rectangle = nested.Children.AddRectangle(Rect(0, 0, 20, 10));
        rectangle.Transform = "translate(10pt 5pt)";
        var connector = group.Children.AddConnector(new OfficePoint(20, 30), new OfficePoint(100, 80), OdgConnectorKind.Curve);
        connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(20, 30), OfficePathCommand.CubicBezierTo(40, 0, 70, 100, 100, 80) });
        group.TransformChildren("scale(2 1) translate(50pt 20pt)");
        Assert.Null(group.Transform); Assert.Null(nested.Transform); Assert.Null(connector.Transform);
        AssertClose(new OfficePoint(90, 50), connector.ConnectorRouteCommands[0].Point);
        AssertClose(new OfficePoint(250, 100), connector.ConnectorRouteCommands[1].Point);
        var read = FlatRoundTrip(document);
        AssertClose(new OfficePoint(250, 100), read.Pages[0].Shapes[0].Children[1].ConnectorRouteCommands[1].Point);
        Assert.False(read.Pages[0].ToDrawing().Report.HasSkippedOrUnsupported);
        Assert.Throws<NotSupportedException>(() => group.Transform = "translate(1cm 1cm)");
    }

    [Fact]
    public void GroupTransformRejectsUnresolvedAttachmentsBeforeChangingAnyChild() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var group = page.Shapes.AddGroup();
        var first = group.Children.AddRectangle(Rect(0, 0, 20, 10)); var last = group.Children.AddEllipse(Rect(100, 40, 20, 10));
        group.Children.AddConnector(first.AddGluePoint(), last.AddGluePoint());
        string before = group.ToXml().ToString();
        Assert.Throws<NotSupportedException>(() => group.TransformChildren("scale(2 1)"));
        Assert.Equal(before, group.ToXml().ToString());
    }

    [Fact]
    public void AutomaticAttachmentPrefersNativePathEdgeOverStaleSavedEndpointAttributes() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var target = page.Shapes.AddEllipse(Rect(100, 40, 20, 20));
        var connector = page.Shapes.AddConnector(new OfficePoint(0, 0), new OfficePoint(100, 50)); connector.AttachEndToShape(target);
        connector.Element.SetAttributeValue(OdfNamespaces.Svg + "x2", "100pt"); connector.Element.SetAttributeValue(OdfNamespaces.Svg + "y2", "50pt");
        connector.Element.SetAttributeValue(OdfNamespaces.Svg + "viewBox", "0 0 3882 1412");
        connector.Element.SetAttributeValue(OdfNamespaces.Svg + "d", "M0 0L3881 1411");
        AssertClose(new OfficePoint(110, 40), new OfficePoint(connector.X2.ToPoints(), connector.Y2.ToPoints()));
        AssertClose(new OfficePoint(110, 40), connector.ConnectorRouteCommands[1].Point);
    }

    [Fact]
    public void ProjectsNativeCircleSpellingButReportsArcAndSectorGeometry() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var circle = page.Shapes.AddEllipse(Rect(20, 30, 40, 40)); circle.Element.Name = OdfNamespaces.Draw + "circle";
        Assert.Equal(OfficeShapeKind.Ellipse, Assert.Single(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingShape>()).Shape.Kind);
        circle.Element.SetAttributeValue(OdfNamespaces.Draw + "kind", "section");
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        circle.Element.Name = OdfNamespaces.Draw + "ellipse";
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    private static OdgDocument LoadFixture(string name) => OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", name));
    private static OdfRect Rect(double x, double y, double width, double height) => new(OdfLength.Points(x), OdfLength.Points(y), OdfLength.Points(width), OdfLength.Points(height));
    private static OdgDocument FlatRoundTrip(OdgDocument document) { using var stream = new MemoryStream(); document.SaveFlatXml(stream); stream.Position = 0; return OdgDocument.LoadFlatXml(stream); }
    private static void AssertClose(OfficePoint expected, OfficePoint actual) { Assert.InRange(Math.Abs(expected.X - actual.X), 0, 0.1); Assert.InRange(Math.Abs(expected.Y - actual.Y), 0, 0.1); }
}
