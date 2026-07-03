using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioShapeLayoutTests {
        [Fact]
        public void ShapeLayoutCellsSaveLoadAndCanBeCleared() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string clearedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Layout", 10, 7);
            VisioShape source = page.AddRectangle(2, 4, 1.6, 0.8, "Source");
            VisioShape target = page.AddRectangle(7, 4, 1.6, 0.8, "Target");
            source.PlacementStyle = VisioPlacementStyle.HierarchyLeftToRightMiddle;
            source.PlacementFlip = VisioPlacementFlip.Horizontal | VisioPlacementFlip.Rotate90;
            source.PlowCode = VisioShapePlowCode.Always;
            source.AllowPlacementOnTop = false;
            source.AllowHorizontalConnectorRoutingThrough = true;
            source.AllowVerticalConnectorRoutingThrough = false;
            source.CanSplitShapes = true;
            source.CanBeSplit = false;
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic);
            connector.Label = "handoff";
            connector.RouteStyle = VisioPageRouteStyle.SimpleHorizontalVertical;
            connector.RouteAppearance = VisioLineRouteExtension.Curved;
            connector.LineJumpStyle = VisioLineJumpStyle.Square;
            connector.LineJumpCode = VisioConnectorLineJumpCode.Always;
            connector.HorizontalJumpDirection = VisioHorizontalLineJumpDirection.Up;
            connector.VerticalJumpDirection = VisioVerticalLineJumpDirection.Right;
            connector.RerouteBehavior = VisioConnectorRerouteBehavior.OnCrossover;

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertShapeLayoutXml(filePath, "Source", "23", "5", "2", "0", "1", "0", "1", "0");
            AssertConnectorLayoutXml(filePath, "handoff", "21", "2", "3", "2", "1", "2", "3");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages.Single(current => current.Name == "Layout");
            VisioShape loadedSource = loadedPage.Shapes.Single(shape => shape.Text == "Source");
            VisioConnector loadedConnector = loadedPage.Connectors.Single(current => current.Label == "handoff");
            Assert.Equal(VisioPlacementStyle.HierarchyLeftToRightMiddle, loadedSource.PlacementStyle);
            Assert.Equal(VisioPlacementFlip.Horizontal | VisioPlacementFlip.Rotate90, loadedSource.PlacementFlip);
            Assert.Equal(VisioShapePlowCode.Always, loadedSource.PlowCode);
            Assert.False(loadedSource.AllowPlacementOnTop);
            Assert.True(loadedSource.AllowHorizontalConnectorRoutingThrough);
            Assert.False(loadedSource.AllowVerticalConnectorRoutingThrough);
            Assert.True(loadedSource.CanSplitShapes);
            Assert.False(loadedSource.CanBeSplit);
            Assert.Equal(VisioPageRouteStyle.SimpleHorizontalVertical, loadedConnector.RouteStyle);
            Assert.Equal(VisioLineRouteExtension.Curved, loadedConnector.RouteAppearance);
            Assert.Equal(VisioLineJumpStyle.Square, loadedConnector.LineJumpStyle);
            Assert.Equal(VisioConnectorLineJumpCode.Always, loadedConnector.LineJumpCode);
            Assert.Equal(VisioHorizontalLineJumpDirection.Up, loadedConnector.HorizontalJumpDirection);
            Assert.Equal(VisioVerticalLineJumpDirection.Right, loadedConnector.VerticalJumpDirection);
            Assert.Equal(VisioConnectorRerouteBehavior.OnCrossover, loadedConnector.RerouteBehavior);

            loadedSource.PlacementStyle = VisioPlacementStyle.TopToBottom;
            loadedSource.PlacementFlip = VisioPlacementFlip.None;
            loadedSource.PlowCode = VisioShapePlowCode.Never;
            loadedSource.AllowPlacementOnTop = true;
            loadedSource.AllowHorizontalConnectorRoutingThrough = false;
            loadedSource.AllowVerticalConnectorRoutingThrough = true;
            loadedSource.CanSplitShapes = false;
            loadedSource.CanBeSplit = true;
            loadedConnector.RouteStyle = VisioPageRouteStyle.FlowchartLeftToRight;
            loadedConnector.RouteAppearance = VisioLineRouteExtension.Straight;
            loadedConnector.LineJumpStyle = VisioLineJumpStyle.Gap;
            loadedConnector.LineJumpCode = VisioConnectorLineJumpCode.Never;
            loadedConnector.HorizontalJumpDirection = VisioHorizontalLineJumpDirection.Down;
            loadedConnector.VerticalJumpDirection = VisioVerticalLineJumpDirection.Left;
            loadedConnector.RerouteBehavior = VisioConnectorRerouteBehavior.Never;
            loaded.Save(roundTripPath);

            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertShapeLayoutXml(roundTripPath, "Source", "1", "8", "1", "1", "0", "1", "0", "1");
            AssertConnectorLayoutXml(roundTripPath, "handoff", "6", "1", "2", "1", "2", "1", "2");

            VisioDocument cleared = VisioDocument.Load(roundTripPath);
            VisioPage clearedPage = cleared.Pages.Single(current => current.Name == "Layout");
            clearedPage.Shapes.Single(shape => shape.Text == "Source").ClearLayoutPolicy();
            clearedPage.Connectors.Single(current => current.Label == "handoff").ClearRoutingPolicy();
            cleared.Save(clearedPath);

            Assert.Empty(VisioValidator.Validate(clearedPath));
            AssertShapeLayoutXml(clearedPath, "Source", null, null, null, null, null, null, null, null);
            AssertConnectorLayoutXml(clearedPath, "handoff", null, null, null, null, null, null, null);
        }

        private static void AssertShapeLayoutXml(
            string filePath,
            string text,
            string? expectedPlaceStyle,
            string? expectedPlaceFlip,
            string? expectedPlowCode,
            string? expectedPermeablePlace,
            string? expectedPermeableX,
            string? expectedPermeableY,
            string? expectedShapeSplit,
            string? expectedShapeSplittable) {
            XElement shape = FindShapeByText(filePath, text);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";

            AssertOptionalCellValue(shape, ns, "ShapePlaceStyle", expectedPlaceStyle);
            AssertOptionalCellValue(shape, ns, "ShapePlaceFlip", expectedPlaceFlip);
            AssertOptionalCellValue(shape, ns, "ShapePlowCode", expectedPlowCode);
            AssertOptionalCellValue(shape, ns, "ShapePermeablePlace", expectedPermeablePlace, expectedPermeablePlace == null ? null : "BOOL");
            AssertOptionalCellValue(shape, ns, "ShapePermeableX", expectedPermeableX, expectedPermeableX == null ? null : "BOOL");
            AssertOptionalCellValue(shape, ns, "ShapePermeableY", expectedPermeableY, expectedPermeableY == null ? null : "BOOL");
            AssertOptionalCellValue(shape, ns, "ShapeSplit", expectedShapeSplit);
            AssertOptionalCellValue(shape, ns, "ShapeSplittable", expectedShapeSplittable);
        }

        private static void AssertConnectorLayoutXml(
            string filePath,
            string text,
            string? expectedRouteStyle,
            string? expectedRouteAppearance,
            string? expectedJumpStyle,
            string? expectedJumpCode,
            string? expectedJumpDirX,
            string? expectedJumpDirY,
            string? expectedFixedCode) {
            XElement connector = FindShapeByText(filePath, text);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";

            AssertOptionalCellValue(connector, ns, "ShapeRouteStyle", expectedRouteStyle);
            AssertOptionalCellValue(connector, ns, "ConLineRouteExt", expectedRouteAppearance);
            AssertOptionalCellValue(connector, ns, "ConLineJumpStyle", expectedJumpStyle);
            AssertOptionalCellValue(connector, ns, "ConLineJumpCode", expectedJumpCode);
            AssertOptionalCellValue(connector, ns, "ConLineJumpDirX", expectedJumpDirX);
            AssertOptionalCellValue(connector, ns, "ConLineJumpDirY", expectedJumpDirY);
            AssertOptionalCellValue(connector, ns, "ConFixedCode", expectedFixedCode);
        }

        private static XElement FindShapeByText(string filePath, string text) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            return page.Root!.Element(ns + "Shapes")!.Elements(ns + "Shape")
                .Single(shape => string.Equals(shape.Element(ns + "Text")?.Value, text, StringComparison.Ordinal));
        }

        private static void AssertOptionalCellValue(XElement shape, XNamespace ns, string name, string? expectedValue, string? expectedUnit = null) {
            XElement[] cells = shape.Elements(ns + "Cell")
                .Where(current => (string?)current.Attribute("N") == name)
                .ToArray();
            if (expectedValue == null) {
                Assert.Empty(cells);
                return;
            }

            XElement cell = Assert.Single(cells);
            Assert.Equal(expectedValue, cell.Attribute("V")!.Value);
            if (expectedUnit != null) {
                Assert.Equal(expectedUnit, cell.Attribute("U")?.Value);
            }
        }

        private static XDocument ReadXml(ZipArchive archive, string entryName) {
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}
