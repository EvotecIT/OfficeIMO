using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgConstrainedRoutingTests {
    [Theory]
    [InlineData(OdgGluePointEscapeDirection.Right, OdgGluePointEscapeDirection.Left)]
    [InlineData(OdgGluePointEscapeDirection.Left, OdgGluePointEscapeDirection.Right)]
    [InlineData(OdgGluePointEscapeDirection.Up, OdgGluePointEscapeDirection.Down)]
    [InlineData(OdgGluePointEscapeDirection.Down, OdgGluePointEscapeDirection.Up)]
    [InlineData(OdgGluePointEscapeDirection.Horizontal, OdgGluePointEscapeDirection.Vertical)]
    [InlineData(OdgGluePointEscapeDirection.Vertical, OdgGluePointEscapeDirection.Horizontal)]
    [InlineData(OdgGluePointEscapeDirection.Auto, OdgGluePointEscapeDirection.Auto)]
    public void SavesRoutesThatHonorBothDeparturesAndAvoidAttachedAndSuppliedBounds(OdgGluePointEscapeDirection departure, OdgGluePointEscapeDirection arrival) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(30, 80, 40, 40));
        var end = page.Shapes.AddRectangle(Rect(300, 160, 40, 40));
        var obstacle = page.Shapes.AddRectangle(Rect(160, 110, 50, 60));
        var a = start.AddGluePoint(OdgGluePointAlignment.Right); a.EscapeDirection = departure;
        var b = end.AddGluePoint(OdgGluePointAlignment.Left); b.EscapeDirection = arrival;
        var connector = page.Shapes.AddConnector(a, b); var other = page.Shapes.AddConnector(a, b);
        string aBefore = a.Element.ToString(), bBefore = b.Element.ToString(), otherBefore = other.ToXml().ToString();
        connector.RouteOrthogonalAroundShapes(new[] { obstacle }, respectEscapeDirections: true, padding: 4);
        foreach (var read in RoundTrips(document)) {
            var saved = read.Pages[0].Shapes.First(shape => shape.IsConnector);
            var route = PageRoute(saved); AssertRoute(route, departure, arrival);
            AssertAvoids(route, 155.9, 105.9, 214.1, 174.1);
            AssertAvoids(route, 25.9, 75.9, 74.1, 124.1, skipFirst: true);
            AssertAvoids(route, 295.9, 155.9, 344.1, 204.1, skipLast: true);
            AssertClose(new OfficePoint(70, 100), route[0]); AssertClose(new OfficePoint(300, 180), route.Last());
            Assert.Equal(OdgConnectorKind.Standard, saved.ConnectorKind);
            Assert.False(read.Pages[0].ToDrawing().Report.HasSkippedOrUnsupported);
        }
        string beforeRepeat = connector.ToXml().ToString();
        connector.RouteOrthogonalAroundShapes(new[] { obstacle }, respectEscapeDirections: true, padding: 4);
        Assert.Equal(beforeRepeat, connector.ToXml().ToString());
        Assert.Equal(aBefore, a.Element.ToString()); Assert.Equal(bBefore, b.Element.ToString()); Assert.Equal(otherBefore, other.ToXml().ToString());
        end.Bounds = Rect(300, 200, 40, 40);
        connector.RouteOrthogonalAroundShapes(new[] { obstacle }, respectEscapeDirections: true, padding: 4);
        AssertClose(new OfficePoint(300, 220), PageRoute(connector).Last());
        AssertRoute(PageRoute(connector), departure, arrival);
    }

    [Fact]
    public void SameShapeLoopExitsBothSidesAndKeepsItsMiddleOutsideTheShape() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddRectangle(Rect(140, 60, 80, 80));
        var a = shape.AddGluePoint(OdgGluePointAlignment.Right); a.EscapeDirection = OdgGluePointEscapeDirection.Right;
        var b = shape.AddGluePoint(OdgGluePointAlignment.Left); b.EscapeDirection = OdgGluePointEscapeDirection.Left;
        var connector = page.Shapes.AddConnector(a, b);
        connector.RouteOrthogonalAroundShapes(Array.Empty<OdgShape>(), respectEscapeDirections: true);
        foreach (var read in RoundTrips(document)) {
            var route = PageRoute(read.Pages[0].Shapes[1]);
            AssertRoute(route, OdgGluePointEscapeDirection.Right, OdgGluePointEscapeDirection.Left);
            AssertAvoids(route, 133.9, 53.9, 226.1, 146.1, skipFirst: true, skipLast: true);
            Assert.True(route.Length >= 6);
        }
    }

    [Fact]
    public void InteriorGluePointsExitTheirAttachedBoundsBeforeTurning() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(30, 30, 120, 120)); var end = page.Shapes.AddRectangle(Rect(260, 30, 120, 120));
        var a = start.AddGluePoint(); a.EscapeDirection = OdgGluePointEscapeDirection.Down;
        var b = end.AddGluePoint(); b.EscapeDirection = OdgGluePointEscapeDirection.Down;
        var connector = page.Shapes.AddConnector(a, b);
        connector.RouteOrthogonalAroundShapes(Array.Empty<OdgShape>(), respectEscapeDirections: true);
        var route = PageRoute(connector); AssertRoute(route, OdgGluePointEscapeDirection.Down, OdgGluePointEscapeDirection.Down);
        Assert.True(route[1].Y > 156.1); Assert.True(route[route.Length - 2].Y > 156.1);
        AssertAvoids(route, 23.9, 23.9, 156.1, 156.1, skipFirst: true);
        AssertAvoids(route, 253.9, 23.9, 386.1, 156.1, skipLast: true);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PartiallyAttachedEndpointsHonorTheAttachedConstraint(bool attachedStart) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddRectangle(Rect(30, 80, 40, 40));
        var point = shape.AddGluePoint(OdgGluePointAlignment.Right); point.EscapeDirection = OdgGluePointEscapeDirection.Down;
        var connector = attachedStart ? page.Shapes.AddConnector(new OfficePoint(70, 100), new OfficePoint(300, 180))
            : page.Shapes.AddConnector(new OfficePoint(300, 180), new OfficePoint(70, 100));
        if (attachedStart) connector.AttachStart(point); else connector.AttachEnd(point);
        connector.RouteOrthogonalAroundShapes(Array.Empty<OdgShape>(), respectEscapeDirections: true);
        foreach (var read in RoundTrips(document)) {
            var route = PageRoute(read.Pages[0].Shapes[1]);
            AssertRoute(route, attachedStart ? OdgGluePointEscapeDirection.Down : OdgGluePointEscapeDirection.Auto,
                attachedStart ? OdgGluePointEscapeDirection.Auto : OdgGluePointEscapeDirection.Down);
            AssertAvoids(route, 23.9, 73.9, 76.1, 126.1, skipFirst: attachedStart, skipLast: !attachedStart);
        }
    }

    [Theory]
    [InlineData("translate(20.17pt 30.23pt)")]
    [InlineData("scale(2 3)")]
    [InlineData("rotate(1.5707963267948966)")]
    [InlineData("rotate(0.7853981633974483)")]
    [InlineData("matrix(1 0.3 0.2 1 20pt 30pt)")]
    public void RoutesInPageCoordinatesThroughAffineConnectorTransforms(string transform) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(30, 80, 40, 40)); var end = page.Shapes.AddRectangle(Rect(300, 160, 40, 40));
        var a = start.AddGluePoint(OdgGluePointAlignment.Right); a.EscapeDirection = OdgGluePointEscapeDirection.Down;
        var b = end.AddGluePoint(OdgGluePointAlignment.Left); b.EscapeDirection = OdgGluePointEscapeDirection.Down;
        var connector = page.Shapes.AddConnector(a, b); connector.Transform = transform;
        var obstacle = page.Shapes.AddRectangle(Rect(160, 110, 50, 60));
        connector.RouteOrthogonalAroundShapes(new[] { obstacle }, respectEscapeDirections: true, padding: 4);
        foreach (var read in RoundTrips(document)) {
            var saved = read.Pages[0].Shapes[2]; var route = PageRoute(saved);
            Assert.Equal(transform, saved.Transform); AssertRoute(route, OdgGluePointEscapeDirection.Down, OdgGluePointEscapeDirection.Down);
            AssertClose(new OfficePoint(70, 100), route[0]); AssertClose(new OfficePoint(300, 180), route.Last());
            AssertAvoids(route, 155.9, 105.9, 214.1, 174.1);
            Assert.False(read.Pages[0].ToDrawing().Report.HasSkippedOrUnsupported);
        }
    }

    [Fact]
    public void RoutesAGroupedConnectorInPageSpaceWithoutMovingItsSiblings() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(30, 80, 40, 40)); var end = page.Shapes.AddRectangle(Rect(300, 160, 40, 40));
        var a = start.AddGluePoint(OdgGluePointAlignment.Right); a.EscapeDirection = OdgGluePointEscapeDirection.Down;
        var b = end.AddGluePoint(OdgGluePointAlignment.Left); b.EscapeDirection = OdgGluePointEscapeDirection.Down;
        var group = page.Shapes.AddGroup();
        var connector = group.Children.AddConnector(a, b);
        connector.Transform = "matrix(1 0.3 0.2 1 20pt 30pt)";
        var sibling = group.Children.AddRectangle(Rect(400, 10, 10, 10)); string before = sibling.ToXml().ToString();
        connector.RouteOrthogonalAroundShapes(Array.Empty<OdgShape>(), respectEscapeDirections: true);
        foreach (var read in RoundTrips(document)) {
            var savedGroup = read.Pages[0].Shapes[2]; var route = PageRoute(savedGroup.Children[0]);
            Assert.Null(savedGroup.Transform); Assert.Equal(before, savedGroup.Children[1].ToXml().ToString());
            AssertRoute(route, OdgGluePointEscapeDirection.Down, OdgGluePointEscapeDirection.Down);
            AssertClose(new OfficePoint(70, 100), route[0]); AssertClose(new OfficePoint(300, 180), route.Last());
        }
    }

    [Fact]
    public void FreeEndpointsAllowEitherAxisWhileStillAvoidingObstacles() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(20, 100), new OfficePoint(300, 100));
        var obstacle = page.Shapes.AddRectangle(Rect(130, 70, 70, 60));
        connector.RouteOrthogonalAroundShapes(new[] { obstacle }, respectEscapeDirections: true);
        foreach (var read in RoundTrips(document)) {
            var saved = read.Pages[0].Shapes[0]; var route = PageRoute(saved);
            Assert.Null(saved.StartShapeId); Assert.Null(saved.EndShapeId);
            AssertRoute(route, OdgGluePointEscapeDirection.Auto, OdgGluePointEscapeDirection.Auto);
            AssertAvoids(route, 123.9, 63.9, 206.1, 136.1);
        }
    }

    [Theory]
    [InlineData("blocked")]
    [InlineData("missing-point")]
    [InlineData("invalid-direction")]
    [InlineData("singular")]
    [InlineData("native-grid")]
    public void UnresolvableConstraintsLeaveEveryPackagePartUnchanged(string failure) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(30, 80, 40, 40)); var end = page.Shapes.AddRectangle(Rect(300, 160, 40, 40));
        var a = start.AddGluePoint(OdgGluePointAlignment.Right); a.EscapeDirection = OdgGluePointEscapeDirection.Down;
        var b = end.AddGluePoint(OdgGluePointAlignment.Left); b.EscapeDirection = OdgGluePointEscapeDirection.Down;
        var connector = page.Shapes.AddConnector(a, b);
        connector.RouteOrthogonal(offset: 20);
        IEnumerable<OdgShape> obstacles = Array.Empty<OdgShape>();
        switch (failure) {
            case "blocked": obstacles = new[] { page.Shapes.AddRectangle(Rect(-100, -100, 700, 600)) }; break;
            case "missing-point": a.Element.Remove(); break;
            case "invalid-direction": a.Element.SetAttributeValue(OdfNamespaces.Draw + "escape-direction", "sideways"); break;
            case "singular": connector.Transform = "scale(0 1)"; break;
            case "native-grid": connector.Transform = "scale(1000 1000)"; break;
        }
        var before = document.Package.Entries.ToDictionary(entry => entry.Name, entry => entry.GetBytesForSave());
        if (failure == "invalid-direction") Assert.Throws<InvalidDataException>(() => connector.RouteOrthogonalAroundShapes(obstacles, true));
        else Assert.Throws<NotSupportedException>(() => connector.RouteOrthogonalAroundShapes(obstacles, true));
        var after = document.Package.Entries.ToDictionary(entry => entry.Name, entry => entry.GetBytesForSave());
        Assert.Equal(before.Keys.OrderBy(key => key), after.Keys.OrderBy(key => key));
        foreach (string key in before.Keys) Assert.Equal(before[key], after[key]);
    }

    [Fact]
    public void ExistingOverloadRetainsItsRouteContractWhenConstraintsAreDisabled() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(30, 80, 40, 40)); var end = page.Shapes.AddRectangle(Rect(300, 160, 40, 40));
        var a = start.AddGluePoint(OdgGluePointAlignment.Right); a.EscapeDirection = OdgGluePointEscapeDirection.Left;
        var b = end.AddGluePoint(OdgGluePointAlignment.Left); b.EscapeDirection = OdgGluePointEscapeDirection.Right;
        var connector = page.Shapes.AddConnector(a, b); var obstacle = page.Shapes.AddRectangle(Rect(160, 110, 50, 60));
        connector.RouteOrthogonalAroundShapes(page.Shapes, 4, 12); string oldCall = connector.ToXml().ToString();
        connector.RouteOrthogonalAroundShapes(page.Shapes, false, 4, 12); Assert.Equal(oldCall, connector.ToXml().ToString());
        connector.RouteOrthogonalAroundShapes(page.Shapes, true, 4, 12); Assert.NotEqual(oldCall, connector.ToXml().ToString());
        AssertRoute(PageRoute(connector), OdgGluePointEscapeDirection.Left, OdgGluePointEscapeDirection.Right);
    }

    [Fact]
    public void EditsIndependentProducerConstraintsWithoutChangingGlueOrOtherShapes() {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-routed-glue.odg"));
        var page = document.Pages[0]; int index = page.Shapes.ToList().FindIndex(shape => shape.IsConnector); var connector = page.Shapes[index];
        var target = page.Shapes.Single(shape => shape.XmlId == connector.EndShapeId);
        target.GluePoints.Single(point => point.Id == 4).EscapeDirection = OdgGluePointEscapeDirection.Down;
        var other = page.Shapes.Where((shape, i) => i != index).Select(shape => shape.ToXml().ToString()).ToArray();
        connector.RouteOrthogonalAroundShapes(page.Shapes, respectEscapeDirections: true);
        var read = OdgDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(other, read.Pages[0].Shapes.Where((shape, i) => i != index).Select(shape => shape.ToXml().ToString()).ToArray());
        var saved = read.Pages[0].Shapes.First(shape => shape.IsConnector);
        string startId = (string)saved.Element.Attribute(OdfNamespaces.Draw + "start-glue-point")!;
        string endId = (string)saved.Element.Attribute(OdfNamespaces.Draw + "end-glue-point")!;
        var start = read.Pages[0].Shapes.Single(shape => shape.XmlId == saved.StartShapeId).GluePoints.FirstOrDefault(point => point.Id.ToString() == startId);
        var end = read.Pages[0].Shapes.Single(shape => shape.XmlId == saved.EndShapeId).GluePoints.Single(point => point.Id.ToString() == endId);
        AssertRoute(PageRoute(saved), start?.EscapeDirection ?? OdgGluePointEscapeDirection.Auto, end.EscapeDirection);
    }

    private static OfficePoint[] PageRoute(OdgShape connector) => connector.ConnectorRouteCommands.Select(command => connector.PageTransform.TransformPoint(command.Point)).ToArray();
    private static void AssertRoute(OfficePoint[] route, OdgGluePointEscapeDirection start, OdgGluePointEscapeDirection end) {
        Assert.True(route.Length >= 2);
        for (int i = 1; i < route.Length; i++) {
            double dx = Math.Abs(route[i].X - route[i - 1].X), dy = Math.Abs(route[i].Y - route[i - 1].Y);
            Assert.True(Math.Min(dx, dy) <= 0.1); Assert.True(Math.Max(dx, dy) > 0.1);
        }
        AssertDeparture(route[0], route[1], start); AssertDeparture(route.Last(), route[route.Length - 2], end);
    }
    private static void AssertDeparture(OfficePoint from, OfficePoint to, OdgGluePointEscapeDirection direction) {
        double dx = to.X - from.X, dy = to.Y - from.Y;
        Assert.True(direction switch {
            OdgGluePointEscapeDirection.Left => dx < 0 && Math.Abs(dy) <= 0.1,
            OdgGluePointEscapeDirection.Right => dx > 0 && Math.Abs(dy) <= 0.1,
            OdgGluePointEscapeDirection.Up => dy < 0 && Math.Abs(dx) <= 0.1,
            OdgGluePointEscapeDirection.Down => dy > 0 && Math.Abs(dx) <= 0.1,
            OdgGluePointEscapeDirection.Horizontal => Math.Abs(dy) <= 0.1,
            OdgGluePointEscapeDirection.Vertical => Math.Abs(dx) <= 0.1,
            _ => true
        });
    }
    private static void AssertAvoids(OfficePoint[] route, double left, double top, double right, double bottom, bool skipFirst = false, bool skipLast = false) {
        for (int i = 1; i < route.Length; i++) {
            if ((i == 1 && skipFirst) || (i == route.Length - 1 && skipLast)) continue;
            Assert.False(OfficeGeometry.SegmentIntersectsRectangle((route[i - 1].X, route[i - 1].Y), (route[i].X, route[i].Y), left, top, right, bottom));
        }
    }
    private static void AssertClose(OfficePoint expected, OfficePoint actual) {
        Assert.InRange(Math.Abs(expected.X - actual.X), 0, 0.1); Assert.InRange(Math.Abs(expected.Y - actual.Y), 0, 0.1);
    }
    private static IEnumerable<OdgDocument> RoundTrips(OdgDocument document) {
        yield return document; yield return OdgDocument.Load(new MemoryStream(document.ToBytes()));
        using var stream = new MemoryStream(); document.SaveFlatXml(stream); stream.Position = 0; yield return OdgDocument.LoadFlatXml(stream);
    }
    private static OdfRect Rect(double x, double y, double w, double h) => new(OdfLength.Points(x), OdfLength.Points(y), OdfLength.Points(w), OdfLength.Points(h));
}
