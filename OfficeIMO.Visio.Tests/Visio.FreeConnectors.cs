using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Reader.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioFreeConnectorTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FreeAndPartialEndpointsEditAndRoundTripWithoutInventedShapes(bool package) {
        var document = VisioDocument.Create(); var page = document.AddPage("Free routes");
        page.DefaultUnit = VisioMeasurementUnit.Centimeters;
        var connector = page.AddConnector("route", new OfficePoint(2.54, 5.08), new OfficePoint(12.7, 7.62));
        AssertPoint(1, 2, connector.StartPoint); AssertPoint(5, 3, connector.EndPoint);
        Assert.Empty(page.Shapes); Assert.Null(connector.From); Assert.Null(connector.To);
        var shape = new VisioShape("target", 5, 3, 2, 2, "Target"); page.Shapes.Add(shape);
        page.ReconnectConnectorEnd(connector, shape, VisioSide.Left);
        shape.PinY += 1;
        AssertPoint(1, 2, connector.StartPoint); AssertPoint(4, 4, connector.EndPoint);
        connector.StartPoint = new OfficePoint(2, 2);
        connector.RouteOrthogonal();
        var saved = RoundTrip(document, package); var copy = Assert.Single(saved.Pages[0].Connectors);
        Assert.Null(copy.From); Assert.Equal("target", copy.To?.Id);
        AssertPoint(2, 2, copy.StartPoint); AssertPoint(4, 4, copy.EndPoint);
        Assert.Single(saved.Pages[0].Shapes); Assert.Equal(2, copy.Waypoints.Count);
        Assert.Single(LegacyConnects(saved));
        var projected = saved.ToOfficeDocumentModel();
        Assert.Contains(projected.Blocks, block => block.Text == "(2, 2) in -> target");
        Assert.Contains("(2, 2) in -> target", saved.ToOfficeDocumentReadResult().Chunks[0].Text);
        var snapshot = Assert.Single(saved.CreateInspectionSnapshot().Pages[0].Connectors);
        Assert.Null(snapshot.FromId); AssertPoint(2, 2, snapshot.StartPoint);
        Assert.Single(saved.Pages[0].IncomingConnectors(saved.Pages[0].Shapes[0]));
        Assert.Empty(saved.Pages[0].ConnectedShapes(saved.Pages[0].Shapes[0]));
        copy.EndPoint = copy.EndPoint;
        Assert.Null(copy.To); Assert.Empty(LegacyConnects(saved));
        copy.StartPoint = new OfficePoint(-1, 0);
        var detached = Assert.Single(RoundTrip(saved, package).Pages[0].Connectors);
        AssertPoint(-1, 0, detached.StartPoint); AssertPoint(4, 4, detached.EndPoint);
        Assert.Null(detached.From); Assert.Null(detached.To);
        Assert.Throws<ArgumentOutOfRangeException>(() => detached.StartPoint = new OfficePoint(double.NaN, 0));
        Assert.Throws<ArgumentException>(() => detached.RouteSelfLoop());
    }

    [Fact]
    public void DuplicationIncludesFreePageRoutesAndPartialSelectionRoutes() {
        var document = VisioDocument.Create(); var page = document.AddPage("Original");
        var shape = new VisioShape("box", 4, 3, 2, 2, "Box"); page.Shapes.Add(shape);
        var free = page.AddConnector("free", new OfficePoint(1, 1), new OfficePoint(2, 2));
        var partial = page.AddConnector("partial", new OfficePoint(0, 0), new OfficePoint(3, 3));
        page.ReconnectConnectorEnd(partial, shape, VisioSide.Left);
        var duplicate = document.DuplicatePage(page, "Copy");
        Assert.Equal(2, duplicate.Connectors.Count);
        var copiedPartial = Assert.Single(duplicate.Connectors, c => c.To != null);
        Assert.Same(duplicate.Shapes[0], copiedPartial.To); Assert.Null(copiedPartial.From);
        AssertPoint(0, 0, copiedPartial.StartPoint);
        page.DuplicateShapes(new[] { shape }, offsetX: 2, offsetY: 1);
        var selectedCopy = page.Connectors.Last();
        Assert.Equal(3, page.Connectors.Count); AssertPoint(2, 1, selectedCopy.StartPoint); AssertPoint(5, 4, selectedCopy.EndPoint);
        Assert.NotSame(shape, selectedCopy.To);
        Assert.Equal(3, RoundTrip(document, true).Pages[0].Connectors.Count);
    }

    [Fact]
    public void IndependentPartialNurbsConnectorKeepsCurveAndFollowsItsAttachedShape() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "clickhouse-replication.vdx");
        var imported = VisioDocument.LoadLegacyXml(path);
        var page = imported.Value.Pages[0]; var curve = page.Connectors.Single(c => c.Id == "34");
        Assert.Equal(8, page.Connectors.Count); Assert.Null(curve.From); Assert.Equal("3", curve.To?.Id);
        AssertPoint(1.6832729468599, 8.01419082125604, curve.StartPoint);
        AssertPoint(5.06660611202271, 6.96715886639225, curve.EndPoint);
        Assert.Equal(ConnectorKind.Curved, curve.Kind);
        Assert.True(CurvePath(page, "34").Count(c => c == 'L') > 10);
        var unedited = RoundTrip(imported.Value, false);
        Assert.Equal(CurvePath(page, "34"), CurvePath(unedited.Pages[0], "34"));
        curve.To!.PinX += 1;
        AssertPoint(6.06660611202271, 6.96715886639225, curve.EndPoint);
        curve.StartPoint = new OfficePoint(1, 8);
        string editedPath = CurvePath(page, "34");
        Assert.NotEqual(CurvePath(unedited.Pages[0], "34"), editedPath);
        foreach (bool package in new[] { false, true }) {
            var saved = RoundTrip(imported.Value, package);
            Assert.Equal(editedPath, CurvePath(saved.Pages[0], "34"));
            var xml = XDocument.Load(new MemoryStream(saved.ToLegacyXmlResult().Value));
            var native = xml.Descendants(Legacy + "Page").Elements(Legacy + "Shapes").Elements(Legacy + "Shape").Single(s => (string?)s.Attribute("ID") == "34");
            Assert.Equal(3, native.Descendants(Legacy + "NURBSTo").Count());
            Assert.Single(LegacyConnects(saved).Where(c => (string?)c.Attribute("FromSheet") == "34"));
            var reloaded = saved.Pages[0].Connectors.Single(c => c.Id == "34");
            AssertPoint(1, 8, reloaded.StartPoint); AssertPoint(6.06660611202271, 6.96715886639225, reloaded.EndPoint);
        }
    }

    [Fact]
    public void AttachedEndpointUsesNestedShapeCoordinates() {
        var document = VisioDocument.Create(); var page = document.AddPage("Group");
        var group = new VisioShape("group", 5, 5, 4, 4, "") { Angle = Math.PI / 2 };
        var child = new VisioShape("child", 1, 1, 2, 2, ""); group.Children.Add(child); page.Shapes.Add(group);
        var connector = page.AddConnector("route", new OfficePoint(1, 1), new OfficePoint(2, 2));
        page.ReconnectConnectorEnd(connector, child, VisioSide.Left);
        AssertPoint(6, 3, connector.EndPoint);
        group.PinX += 2;
        AssertPoint(8, 3, connector.EndPoint);
        AssertPoint(8, 3, RoundTrip(document, true).Pages[0].Connectors[0].EndPoint);
    }

    [Fact]
    public void AbsoluteNurbsControlCoordinatesScaleWithEditedEndpoints() {
        const string xml = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Pages><Page ID='0'><Shapes><Shape ID='1'><XForm><PinX>2</PinX><PinY>1</PinY><Width>4</Width><Height>2</Height><LocPinX>2</LocPinX><LocPinY>1</LocPinY></XForm><XForm1D><BeginX>0</BeginX><BeginY>0</BeginY><EndX>4</EndX><EndY>0</EndY></XForm1D><Geom IX='0'><NoFill>1</NoFill><MoveTo IX='1'><X>0</X><Y>0</Y></MoveTo><NURBSTo IX='2'><X>4</X><Y>0</Y><A>1</A><B>1</B><C>0</C><D>1</D><E>NURBS(1,3,1,1,1,2,0,1,3,2,0,1)</E></NURBSTo></Geom></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml))).Value;
        var connector = Assert.Single(document.Pages[0].Connectors);
        connector.EndPoint = new OfficePoint(8, 0);
        string path = CurvePath(document.Pages[0], "1");
        Assert.Equal(path, CurvePath(RoundTrip(document, true).Pages[0], "1"));
        var saved = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        Assert.Equal("NURBS(1,3,1,1,2,4,0,1,6,4,0,1)", Assert.Single(saved.Descendants(Legacy + "NURBSTo")).Element(Legacy + "E")!.Value);
        connector.RouteThrough(new VisioConnectorWaypoint(4, 1));
        Assert.Empty(XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value)).Descendants(Legacy + "NURBSTo"));
    }

    [Theory]
    [InlineData(5, 1)]
    [InlineData(6, 5)]
    public void LabelPinsAndRotationUseTheConnectorLocalFrame(double endX, double endY) {
        var document = VisioDocument.Create(); var page = document.AddPage("Labels");
        var connector = page.AddConnector("route", new OfficePoint(1, 1), new OfficePoint(endX, endY));
        connector.Label = "Review"; connector.LabelPlacement = VisioConnectorLabelPlacement.Along(.5);
        var xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        var shape = xml.Descendants(Legacy + "Page").Elements(Legacy + "Shapes").Elements(Legacy + "Shape").Single();
        var transform = shape.Element(Legacy + "XForm")!; var text = shape.Element(Legacy + "TextXForm")!;
        double Read(XElement parent, string name) => double.Parse(parent.Element(Legacy + name)!.Value, CultureInfo.InvariantCulture);
        var frame = new VisioShape("frame") { PinX = Read(transform, "PinX"), PinY = Read(transform, "PinY"), LocPinX = Read(transform, "LocPinX"), LocPinY = Read(transform, "LocPinY"), Angle = Read(transform, "Angle") };
        var point = frame.GetAbsolutePoint(Read(text, "TxtPinX"), Read(text, "TxtPinY"));
        Assert.Equal((1 + endX) / 2, point.X, 8); Assert.Equal((1 + endY) / 2, point.Y, 8);
        Assert.Equal(0, frame.Angle + Read(text, "TxtAngle"), 8);
        foreach (bool package in new[] { false, true }) {
            var loaded = RoundTrip(document, package).Pages[0].Connectors[0];
            Assert.Equal(point.X, loaded.LabelPlacement!.PinX!.Value, 8); Assert.Equal(point.Y, loaded.LabelPlacement.PinY!.Value, 8);
            Assert.Equal(0, loaded.TextStyle!.TextAngle!.Value, 8);
        }
    }

    [Fact]
    public void ReopenedWaypointRouteKeepsItsPagePointsWhenAnEndpointMoves() {
        var document = VisioDocument.Create(); var page = document.AddPage("Route");
        var connector = page.AddConnector("route", new OfficePoint(0, 0), new OfficePoint(4, 4)).RouteThrough(new VisioConnectorWaypoint(0, 4));
        var loaded = RoundTrip(document, true);
        connector.EndPoint = new OfficePoint(8, 4);
        loaded.Pages[0].Connectors[0].EndPoint = new OfficePoint(8, 4);
        Assert.Equal(CurvePath(page, "route"), CurvePath(loaded.Pages[0], "route"));
        Assert.Equal(CurvePath(page, "route"), CurvePath(RoundTrip(loaded, false).Pages[0], "route"));
    }

    [Fact]
    public void DuplicatingAnEditedWaypointListDoesNotRestoreTheOldNativeRoute() {
        var document = VisioDocument.Create(); var page = document.AddPage("Route");
        page.AddConnector("route", new OfficePoint(0, 0), new OfficePoint(4, 4)).RouteThrough(new VisioConnectorWaypoint(0, 4));
        document = RoundTrip(document, true); page = document.Pages[0];
        page.Connectors[0].Waypoints[0] = new VisioConnectorWaypoint(2, 5);
        var duplicate = document.DuplicatePage(page, "Copy");
        Assert.Equal(CurvePath(page, "route"), CurvePath(duplicate, duplicate.Connectors[0].Id));
        var loaded = RoundTrip(document, false);
        Assert.Equal(CurvePath(loaded.Pages[0], "route"), CurvePath(loaded.Pages[1], loaded.Pages[1].Connectors[0].Id));
    }

    private static VisioDocument RoundTrip(VisioDocument document, bool package) => package
        ? VisioDocument.Load(new MemoryStream(document.ToBytes()))
        : VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    private static XElement[] LegacyConnects(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value)).Descendants(Legacy + "Connect").ToArray();
    private static string CurvePath(VisioPage page, string id) => (string)XDocument.Parse(page.ToSvg()).Descendants()
        .Single(e => (string?)e.Attribute("data-visio-connector-id") == id).Elements().First(e => e.Name.LocalName == "path").Attribute("d")!;
    private static void AssertPoint(double x, double y, OfficePoint point) { Assert.Equal(x, point.X, 7); Assert.Equal(y, point.Y, 7); }
}
