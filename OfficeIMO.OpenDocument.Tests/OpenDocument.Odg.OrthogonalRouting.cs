using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgOrthogonalRoutingTests {
    [Theory]
    [InlineData(OdgConnectorRouteOrientation.Auto, true)]
    [InlineData(OdgConnectorRouteOrientation.HorizontalFirst, true)]
    [InlineData(OdgConnectorRouteOrientation.VerticalFirst, false)]
    public void GeneratesAndSavesThreeSegmentRoutes(OdgConnectorRouteOrientation orientation, bool horizontal) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(10, 20), new OfficePoint(130, 80));
        connector.RouteOrthogonal(orientation, 12);
        foreach (var read in RoundTrips(document)) {
            var saved = read.Pages[0].Shapes[0]; Assert.Equal(OdgConnectorKind.Standard, saved.ConnectorKind);
            var route = saved.ConnectorRouteCommands;
            Assert.Equal(4, route.Count); AssertOrthogonal(route);
            if (horizontal) { AssertClose(82, route[1].Point.X); AssertClose(20, route[1].Point.Y); }
            else { AssertClose(10, route[1].Point.X); AssertClose(62, route[1].Point.Y); }
            Assert.False(read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Report.HasSkippedOrUnsupported);
        }
    }

    [Fact]
    public void AvoidsObstacleBoundsAndIgnoresAttachedShapesAcrossRoundTrips() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var start = page.Shapes.AddRectangle(Rect(0, 20, 20, 20));
        var end = page.Shapes.AddRectangle(Rect(120, 20, 20, 20));
        page.Shapes.AddRectangle(Rect(50, 10, 40, 40));
        var connector = page.Shapes.AddConnector(start.AddGluePoint(OdgGluePointAlignment.Right), end.AddGluePoint(OdgGluePointAlignment.Left));
        connector.RouteOrthogonalAroundShapes(page.Shapes, padding: 4);
        foreach (var read in RoundTrips(document)) {
            var saved = read.Pages[0].Shapes.Single(shape => shape.IsConnector);
            var route = saved.ConnectorRouteCommands; AssertOrthogonal(route);
            AssertAvoids(route, 46, 6, 94, 54);
            AssertClose(20, route[0].Point.X); AssertClose(120, route[route.Count - 1].Point.X);
            Assert.False(read.Pages[0].ToDrawing().Report.HasSkippedOrUnsupported);
        }
        // Moving the attachment invalidates the cache; routing resolves the new endpoint before replacing it.
        end.Bounds = Rect(120, 80, 20, 20);
        Assert.Throws<NotSupportedException>(() => connector.ConnectorRouteCommands);
        connector.RouteOrthogonalAroundShapes(page.Shapes, padding: 4);
        AssertClose(90, connector.ConnectorRouteCommands.Last().Point.Y);
        AssertAvoids(connector.ConnectorRouteCommands, 46, 6, 94, 54);
    }

    [Fact]
    public void RoutesInConnectorCoordinatesAroundTransformedObstacle() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var obstacle = page.Shapes.AddRectangle(Rect(20, 20, 20, 40));
        obstacle.Transform = "scale(2 1) translate(10pt 20pt)";
        var connector = page.Shapes.AddConnector(new OfficePoint(0, 40), new OfficePoint(120, 40));
        connector.Transform = "translate(10pt 20pt)";
        connector.RouteOrthogonalAroundShapes(new[] { obstacle }, padding: 0);
        foreach (var read in RoundTrips(document)) {
            var route = read.Pages[0].Shapes[1].ConnectorRouteCommands;
            AssertOrthogonal(route); AssertAvoids(route, 40, 20, 80, 60);
        }
    }

    [Theory]
    [InlineData("blocked")]
    [InlineData("group")]
    [InlineData("foreign-page")]
    [InlineData("null")]
    [InlineData("failed-enumeration")]
    [InlineData("too-many")]
    [InlineData("invalid-padding")]
    [InlineData("invalid-lanes")]
    public void FailedObstacleSearchRetainsExistingCurveAndXml(string failure) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var connector = page.Shapes.AddConnector(new OfficePoint(10, 20), new OfficePoint(100, 80), OdgConnectorKind.Curve);
        connector.SetConnectorRoute(new[] { OfficePathCommand.MoveTo(10, 20), OfficePathCommand.CubicBezierTo(20, 0, 80, 100, 100, 80) });
        var obstacle = page.Shapes.AddRectangle(Rect(-100, -100, 500, 500));
        IEnumerable<OdgShape> input = failure switch {
            "group" => new[] { page.Shapes.AddGroup() },
            "foreign-page" => new[] { document.AddPage().Shapes.AddRectangle(Rect(0, 0, 10, 10)) },
            "null" => new OdgShape[] { null! },
            "failed-enumeration" => FailingObstacles(obstacle),
            "too-many" => Enumerable.Repeat(obstacle, 4097),
            _ => new[] { obstacle }
        };
        string before = connector.ToXml().ToString();
        if (failure == "failed-enumeration") Assert.Throws<InvalidOperationException>(() => connector.RouteOrthogonalAroundShapes(input));
        else if (failure is "blocked" or "group") Assert.Throws<NotSupportedException>(() => connector.RouteOrthogonalAroundShapes(input));
        else Assert.ThrowsAny<ArgumentException>(() => connector.RouteOrthogonalAroundShapes(input,
            padding: failure == "invalid-padding" ? double.NaN : 6, maxLanes: failure == "invalid-lanes" ? 33 : 12));
        Assert.Equal(before, connector.ToXml().ToString());
    }

    [Fact]
    public void InvalidLaneOffsetRetainsRoutingKindAndSavedGeometry() {
        var document = OdgDocument.Create(); var connector = document.AddPage().Shapes.AddConnector(new OfficePoint(0, 0), new OfficePoint(20, 10));
        string before = connector.ToXml().ToString();
        Assert.Throws<ArgumentException>(() => connector.RouteOrthogonal(offset: double.PositiveInfinity));
        Assert.Equal(before, connector.ToXml().ToString());
    }

    [Fact]
    public void ReplacesIndependentProducerRouteWhilePreservingOtherShapes() {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-routed-glue.odg"));
        var connector = document.Pages[0].Shapes.First(shape => shape.IsConnector);
        var other = document.Pages[0].Shapes.Where(shape => !shape.IsConnector).Select(shape => shape.ToXml().ToString()).ToArray();
        connector.RouteOrthogonal(offset: 10);
        var reopened = OdgDocument.Load(new MemoryStream(document.ToBytes()));
        AssertOrthogonal(reopened.Pages[0].Shapes.First(shape => shape.IsConnector).ConnectorRouteCommands);
        Assert.Equal(other, reopened.Pages[0].Shapes.Where(shape => !shape.IsConnector).Select(shape => shape.ToXml().ToString()).ToArray());
    }

    private static IEnumerable<OdgShape> FailingObstacles(OdgShape first) { yield return first; throw new InvalidOperationException("Obstacle source failed."); }
    private static IEnumerable<OdgDocument> RoundTrips(OdgDocument document) {
        yield return document; yield return OdgDocument.Load(new MemoryStream(document.ToBytes()));
        using var stream = new MemoryStream(); document.SaveFlatXml(stream); stream.Position = 0; yield return OdgDocument.LoadFlatXml(stream);
    }
    private static OdfRect Rect(double x, double y, double w, double h) => new(OdfLength.Points(x), OdfLength.Points(y), OdfLength.Points(w), OdfLength.Points(h));
    private static void AssertClose(double expected, double actual) => Assert.InRange(Math.Abs(expected - actual), 0, 0.1);
    private static void AssertOrthogonal(IReadOnlyList<OfficePathCommand> route) {
        for (int i = 1; i < route.Count; i++) {
            Assert.Equal(OfficePathCommandKind.LineTo, route[i].Kind);
            Assert.True(Math.Abs(route[i].Point.X - route[i - 1].Point.X) < 0.1 || Math.Abs(route[i].Point.Y - route[i - 1].Point.Y) < 0.1);
        }
    }
    private static void AssertAvoids(IReadOnlyList<OfficePathCommand> route, double left, double top, double right, double bottom) {
        for (int i = 1; i < route.Count; i++) Assert.False(OfficeGeometry.SegmentIntersectsRectangle(
            (route[i - 1].Point.X, route[i - 1].Point.Y), (route[i].Point.X, route[i].Point.Y), left, top, right, bottom));
    }
}
