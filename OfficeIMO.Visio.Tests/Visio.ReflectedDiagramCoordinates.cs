using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;
using ShapeBounds = OfficeIMO.Visio.VisioShapeBounds;

namespace OfficeIMO.Tests;

public sealed class VisioReflectedDiagramCoordinateTests {
    // Origin coordinates also match the independent outline controls in the rendering tests.
    [Theory]
    [InlineData(false, false, 0, false, 0, 1.7, 2.8, 1.7, 2.8, 3.7, 3.8)]
    [InlineData(true, false, 0, false, 0, 2.3, 2.8, 0.3, 2.8, 2.3, 3.8)]
    [InlineData(false, true, 0, false, 0, 1.7, 3.2, 1.7, 2.2, 3.7, 3.2)]
    [InlineData(true, false, 30, false, 0, 2.359807621135, 2.976794919243, 0.127756813566, 1.976794919243, 2.359807621135, 3.842820323028)]
    [InlineData(false, false, 0, true, 0, 2.3, 5.8, 0.3, 5.8, 2.3, 6.8)]
    [InlineData(false, true, 30, true, 45, 0.702480032933, 3.977937659975, -1.229371619645, 2.494373743481, 0.961299078036, 3.977937659975)]
    public void PagePointsBoundsQueriesAndExplicitGlueUseTheRenderedShapeFrame(
        bool flipX, bool flipY, double angle, bool group, double groupAngle,
        double originX, double originY, double left, double bottom, double right, double top) {
        VisioDocument document = Create(flipX, flipY, angle, group, groupAngle);
        VisioPage page = document.Pages[0];
        VisioShape shape = page.FindShapeById("1")!;
        var point = new VisioConnectionPoint(0, 0, 1, 0);
        shape.ConnectionPoints.Add(point);
        VisioShape target = page.AddRectangle(7, 7, 1, 1, "Target");
        VisioConnector connector = page.AddConnector(shape, target, ConnectorKind.Straight);
        connector.FromConnectionPoint = point;
        foreach (VisioDocument current in Reopenings(document)) {
            VisioShape leaf = current.Pages[0].FindShapeById("1")!;
            var actual = leaf.GetAbsolutePoint(0, 0);
            Assert.Equal(originX, actual.X, 8); Assert.Equal(originY, actual.Y, 8);
            AssertBounds(left, bottom, right, top, leaf.GetShapeBounds());
            AssertBounds(left, bottom, right, top, leaf.GetPageShapeBounds());
            ShapeBounds query = new(left - 0.01, bottom - 0.01, right + 0.01, top + 0.01);
            Assert.Contains(leaf, current.Pages[0].ShapesContainedIn(query));
            VisioConnector saved = Assert.Single(current.Pages[0].Connectors);
            Assert.Equal(originX, saved.StartPoint.X, 8); Assert.Equal(originY, saved.StartPoint.Y, 8);
            leaf.PinX += 1;
            if (leaf.Parent == null) Assert.Equal(originX + 1, saved.StartPoint.X, 8);
            else {
                // Moving the containing group translates the attached endpoint in page coordinates.
                leaf.PinX -= 1; leaf.Parent.PinX += 1;
                Assert.Equal(originX + 1, saved.StartPoint.X, 8);
            }
        }
    }

    [Fact]
    public void PublicAbsolutePointIncludesUnreflectedContainingGroups() {
        var document = VisioDocument.Create(); var page = document.AddPage("Coordinates");
        var group = new VisioShape("group", 5, 5, 4, 4, "") { Angle = Math.PI / 2 };
        var child = new VisioShape("child", 1, 1, 2, 2, "");
        group.Children.Add(child); page.Shapes.Add(group);
        var point = child.GetAbsolutePoint(0, 1);
        Assert.Equal(6, point.X, 8); Assert.Equal(3, point.Y, 8);
        AssertBounds(5, 3, 7, 5, child.GetShapeBounds());
    }

    [Fact]
    public void CachedPageAttachmentKeepsItsPositionThroughRotationAndBothSaveFormats() {
        VisioDocument document = Create(true, false, connector: true);
        VisioShape shape = document.Pages[0].FindShapeById("1")!;
        VisioConnector connector = Assert.Single(document.Pages[0].Connectors);
        Assert.Equal(0.8, connector.StartPoint.X, 8); Assert.Equal(3.3, connector.StartPoint.Y, 8);
        shape.PinX += 1; shape.Angle = Math.PI / 2;
        foreach (VisioDocument current in Reopenings(document)) {
            VisioConnector saved = Assert.Single(current.Pages[0].Connectors);
            Assert.Equal("1", saved.From!.Id); Assert.Null(saved.To);
            Assert.Equal(2.7, saved.StartPoint.X, 8); Assert.Equal(1.8, saved.StartPoint.Y, 8);
            Assert.Equal(6, saved.EndPoint.X, 8); Assert.Equal(3.3, saved.EndPoint.Y, 8);
        }
    }

    [Fact]
    public void ContainerFittingUsesReflectedBoundsAndTheInverseOfItsContainingGroup() {
        VisioDocument document = Create(true, false, group: true, groupAngle: 90);
        VisioPage page = document.Pages[0];
        VisioShape member = page.AddRectangle(6, 3, 2, 1, "Member");
        VisioShape container = page.AddContainer("container", "Container", new[] { member },
            new VisioContainerOptions { Margin = 0.2, HeadingHeight = 0 });
        page.ReparentShape(container, page.FindShapeById("2")!);
        page.RefitContainer(container, new VisioContainerOptions { Margin = 0.2, HeadingHeight = 0 });
        foreach (VisioDocument current in Reopenings(document))
            AssertBounds(4.8, 2.3, 7.2, 3.7, current.Pages[0].FindShapeById("container")!.GetPageShapeBounds());
    }

    [Fact]
    public void ReflectedSelectionAlignmentAndGridPlacementMovePageBounds() {
        VisioDocument document = Create(true, false, group: true, groupAngle: 90);
        VisioPage page = document.Pages[0]; VisioShape leaf = page.FindShapeById("1")!;
        VisioShape target = page.AddRectangle(7, 6, 1, 1, "Target");
        var selection = page.SelectShapes(shape => ReferenceEquals(shape, leaf) || ReferenceEquals(shape, target));
        double expectedLeft = selection.GetShapeBounds().Left;
        selection.Align(VisioHorizontalAlignment.Left);
        Assert.Equal(expectedLeft, leaf.GetPageShapeBounds().Left, 8);
        Assert.Equal(expectedLeft, target.GetPageShapeBounds().Left, 8);
        double expectedTop = selection.GetShapeBounds().Top;
        selection.Align(VisioVerticalAlignment.Top);
        Assert.Equal(expectedTop, leaf.GetPageShapeBounds().Top, 8);
        Assert.Equal(expectedTop, target.GetPageShapeBounds().Top, 8);
        selection.RelayoutAsGrid(columns: 2, horizontalSpacing: 0.5, verticalSpacing: 0.5, routeInternalConnectors: false);
        ShapeBounds first = leaf.GetPageShapeBounds(), second = target.GetPageShapeBounds();
        Assert.Equal(expectedLeft, first.Left, 8); Assert.Equal(expectedTop, first.Top, 8);
        Assert.Equal(first.Right + 0.5, second.Left, 8); Assert.Equal(first.CenterY, second.CenterY, 8);
    }

    [Fact]
    public void NestedDistributionUsesPageCentersAlongBothAxes() {
        VisioDocument document = Create(true, false, group: true, groupAngle: 90);
        VisioPage page = document.Pages[0]; VisioShape leaf = page.FindShapeById("1")!;
        VisioShape left = page.AddRectangle(-0.5, 5, 1, 1, "Left");
        VisioShape right = page.AddRectangle(7, 5, 1, 1, "Right");
        var selection = page.SelectShapes(shape => shape.Parent != null || ReferenceEquals(shape, left) || ReferenceEquals(shape, right));
        double originalY = leaf.GetPageShapeBounds().CenterY;
        selection.DistributeHorizontally();
        Assert.Equal(3.25, leaf.GetPageShapeBounds().CenterX, 8);
        Assert.Equal(originalY, leaf.GetPageShapeBounds().CenterY, 8);
        selection.DistributeVertically();
        double[] centers = selection.Select(shape => shape.GetPageShapeBounds().CenterY).OrderBy(value => value).ToArray();
        Assert.Equal((centers[0] + centers[2]) / 2, centers[1], 8);
    }

    [Fact]
    public void AligningAChildAndItsParentAppliesTheParentFirst() {
        VisioDocument document = Create(false, false, group: true, groupAngle: 90);
        VisioPage page = document.Pages[0];
        var selection = page.SelectShapes(shape => true);
        double expected = selection.GetShapeBounds().Right;
        selection.Align(VisioHorizontalAlignment.Right);
        Assert.All(selection, shape => Assert.Equal(expected, shape.GetPageShapeBounds().Right, 8));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void OverlappingSelectionsLayoutEachDistinctShapeOnce(int operation) {
        (VisioShapeSelection Selection, VisioShape[] Shapes) CreateSelection(bool overlap) {
            var document = VisioDocument.Create(); var page = document.AddPage("Selection");
            VisioShape[] shapes = { page.AddRectangle(2, 2, 1, 1, "A"), page.AddRectangle(4, 4, 1, 1, "B"), page.AddRectangle(8, 8, 1, 1, "C") };
            return (new VisioShapeSelection(overlap ? new[] { shapes[0], shapes[1], shapes[0], shapes[2] } : shapes, page), shapes);
        }
        var control = CreateSelection(false); var overlapping = CreateSelection(true);
        foreach (var selection in new[] { control.Selection, overlapping.Selection }) {
            if (operation == 0) selection.RelayoutAsGrid(routeInternalConnectors: false);
            else if (operation == 1) selection.DistributeHorizontally();
            else selection.DistributeVertically();
        }
        Assert.Equal(4, overlapping.Selection.Count);
        for (int index = 0; index < 3; index++) Assert.Equal(control.Shapes[index].GetBounds(), overlapping.Shapes[index].GetBounds());
    }

    [Fact]
    public void MovingShapesAndContainingGroupsRetainsDynamicReflectionCaches() {
        VisioDocument document = Create(false, false);
        var xml = System.Xml.Linq.XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        System.Xml.Linq.XNamespace ns = "http://schemas.microsoft.com/visio/2003/core";
        xml.Descendants(ns + "FlipX").Single().SetAttributeValue("F", "IF(PinX>3,1,0)");
        document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString()))).Value;
        VisioPage page = document.Pages[0]; VisioShape leaf = page.FindShapeById("1")!;
        page.AddRectangle(6, 3, 1, 1, "Target");
        page.SelectShapes(shape => true).Align(VisioHorizontalAlignment.Right);
        foreach (VisioDocument current in Reopenings(document)) {
            Assert.Equal(6.5, current.Pages[0].FindShapeById("1")!.GetShapeBounds().Right, 8);
            Assert.Contains(current.Pages[0].ToDrawing().Report.FidelityDiagnostics, finding => finding.Code == "VISIO_SHAPE_CACHED_REFLECTION");
        }
        document = Create(false, false, group: true);
        xml = System.Xml.Linq.XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        xml.Descendants(ns + "Shape").Single(shape => (string?)shape.Attribute("ID") == "2").Element(ns + "XForm")!.Element(ns + "FlipX")!.SetAttributeValue("F", "IF(PinX>3,1,0)");
        document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString()))).Value;
        document.Pages[0].FindShapeById("2")!.PinX = 2;
        foreach (VisioDocument current in Reopenings(document)) {
            var origin = current.Pages[0].FindShapeById("1")!.GetAbsolutePoint(0, 0);
            Assert.Equal(0.3, origin.X, 8); Assert.Equal(5.8, origin.Y, 8);
        }
    }

    [Fact]
    public void UnusableReflectionCannotProduceMisleadingPublicCoordinates() {
        VisioDocument document = Create(true, false, flipValue: "NaN");
        VisioShape shape = document.Pages[0].FindShapeById("1")!;
        Assert.Throws<InvalidDataException>(() => shape.GetAbsolutePoint(0, 0));
        Assert.Throws<InvalidDataException>(() => shape.GetPageShapeBounds());
    }

    private static IEnumerable<VisioDocument> Reopenings(VisioDocument document) {
        yield return VisioDocument.Load(new MemoryStream(document.ToBytes()));
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
        yield return document;
    }

    private static void AssertBounds(double left, double bottom, double right, double top, ShapeBounds bounds) {
        Assert.Equal(left, bounds.Left, 8); Assert.Equal(bottom, bounds.Bottom, 8);
        Assert.Equal(right, bounds.Right, 8); Assert.Equal(top, bounds.Top, 8);
    }

    private static VisioDocument Create(bool flipX, bool flipY, double angle = 0,
        bool group = false, double groupAngle = 0, bool connector = false, string? flipValue = null) {
        string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
        string shape = $"<Shape ID='1' NameU='Reflected'><XForm><PinX>2</PinX><PinY>3</PinY><Width>2</Width><Height>1</Height><LocPinX>0.3</LocPinX><LocPinY>0.2</LocPinY><Angle>{Number(angle * Math.PI / 180)}</Angle><FlipX>{flipValue ?? (flipX ? "1" : "0")}</FlipX><FlipY>{(flipY ? 1 : 0)}</FlipY></XForm></Shape>";
        if (group) shape = $"<Shape ID='2' Type='Group'><XForm><PinX>4</PinX><PinY>3</PinY><Width>4</Width><Height>4</Height><LocPinX>0</LocPinX><LocPinY>0</LocPinY><FlipX>1</FlipX><Angle>{Number(groupAngle * Math.PI / 180)}</Angle></XForm><Shapes>{shape}</Shapes></Shape>";
        if (connector) shape += "<Shape ID='9' NameU='Connector'><XForm><PinX>3.4</PinX><PinY>3.3</PinY><Width>5.2</Width><Height>0</Height><LocPinX>2.6</LocPinX><LocPinY>0</LocPinY></XForm><XForm1D><BeginX>0.8</BeginX><BeginY>3.3</BeginY><EndX>6</EndX><EndY>3.3</EndY></XForm1D><Geom IX='0'><NoFill>1</NoFill><MoveTo IX='1'><X>0</X><Y>0</Y></MoveTo><LineTo IX='2'><X>5.2</X><Y>0</Y></LineTo></Geom></Shape>";
        string connections = connector ? "<Connects><Connect FromSheet='9' FromCell='BeginX' ToSheet='1' ToCell='PinX'/></Connects>" : "";
        string xml = $"<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Pages><Page ID='0'><PageSheet><PageProps><PageWidth>8</PageWidth><PageHeight>8</PageHeight></PageProps></PageSheet><Shapes>{shape}</Shapes>{connections}</Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
    }
}
