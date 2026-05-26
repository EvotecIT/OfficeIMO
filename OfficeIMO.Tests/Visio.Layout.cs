using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioLayoutTests {
        [Fact]
        public void ContentBoundsAndFitToContentMoveShapesAndResizePage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Fit", 20, 20);
            VisioShape first = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "first", 4, 5, "First");
            first.Width = 2;
            first.Height = 1;
            VisioShape second = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "second", 9, 2, "Second");
            second.Width = 1;
            second.Height = 2;

            OfficeIMO.Visio.VisioShapeBounds before = page.GetContentBounds();

            Assert.Equal(3, before.Left);
            Assert.Equal(1.5, before.Bottom);
            Assert.Equal(9, before.Right);
            Assert.Equal(5.5, before.Top);

            page.FitToContent(horizontalMargin: 0.5, verticalMargin: 0.25);

            OfficeIMO.Visio.VisioShapeBounds after = page.GetContentBounds();
            Assert.Equal(0.5, after.Left, 6);
            Assert.Equal(0.25, after.Bottom, 6);
            Assert.Equal(6.5, after.Right, 6);
            Assert.Equal(4.25, after.Top, 6);
            Assert.Equal(7.0, page.Width, 6);
            Assert.Equal(4.5, page.Height, 6);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void ShapeBoundsUseLocPinAndRotation() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Transformed", 10, 10);
            VisioShape shape = page.AddRectangle(4, 4, 2, 1, "Rotated");
            shape.LocPinX = 0;
            shape.LocPinY = 0;
            shape.Angle = Math.PI / 2D;

            OfficeIMO.Visio.VisioShapeBounds bounds = shape.GetShapeBounds();

            Assert.Equal(3, bounds.Left, 6);
            Assert.Equal(4, bounds.Bottom, 6);
            Assert.Equal(4, bounds.Right, 6);
            Assert.Equal(6, bounds.Top, 6);
        }

        [Fact]
        public void FitToContentIncludesConnectorWaypointsAndLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("FitConnectors", 20, 20);
            VisioShape source = page.AddRectangle(4, 4, 1, 1, "Source");
            VisioShape target = page.AddRectangle(8, 4, 1, 1, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteThrough(
                    VisioConnectorWaypoint.At(2, 6),
                    VisioConnectorWaypoint.At(11, 6))
                .PlaceLabelAt(12, 6.5, width: 2, height: 0.5);
            connector.Label = "External label";

            OfficeIMO.Visio.VisioShapeBounds shapeOnlyBounds = page.GetContentBounds(includeConnectors: false);
            OfficeIMO.Visio.VisioShapeBounds fullBounds = page.GetContentBounds();

            Assert.Equal(3.5, shapeOnlyBounds.Left);
            Assert.Equal(8.5, shapeOnlyBounds.Right);
            Assert.Equal(2, fullBounds.Left);
            Assert.Equal(13, fullBounds.Right);
            Assert.Equal(6.75, fullBounds.Top);

            page.FitToContent(horizontalMargin: 0.5, verticalMargin: 0.25);

            OfficeIMO.Visio.VisioShapeBounds fittedBounds = page.GetContentBounds();
            Assert.Equal(0.5, fittedBounds.Left, 6);
            Assert.Equal(0.25, fittedBounds.Bottom, 6);
            Assert.Equal(12.0, page.Width, 6);
            Assert.Equal(3.75, page.Height, 6);
            Assert.Equal(0.5, connector.Waypoints[0].X, 6);
            Assert.Equal(2.75, connector.Waypoints[0].Y, 6);
            Assert.Equal(10.5, connector.LabelPlacement!.AbsolutePinX!.Value, 6);
            Assert.Equal(3.25, connector.LabelPlacement.AbsolutePinY!.Value, 6);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void CenterContentMovesConnectorPageCoordinatesWithShapes() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("CenterConnectors", 14, 10);
            VisioShape source = page.AddRectangle(2, 2, 1, 1, "Source");
            VisioShape target = page.AddRectangle(4, 2, 1, 1, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteThrough(VisioConnectorWaypoint.At(3, 5))
                .PlaceLabelAt(3, 5.4, width: 1.4, height: 0.4);
            connector.Label = "Centered";

            page.CenterContent();

            OfficeIMO.Visio.VisioShapeBounds bounds = page.GetContentBounds();
            Assert.Equal(7, bounds.CenterX, 6);
            Assert.Equal(5, bounds.CenterY, 6);
            Assert.Equal(7, connector.Waypoints[0].X, 6);
            Assert.Equal(6.45, connector.Waypoints[0].Y, 6);
            Assert.Equal(7, connector.LabelPlacement!.AbsolutePinX!.Value, 6);
            Assert.Equal(6.85, connector.LabelPlacement.AbsolutePinY!.Value, 6);
        }

        [Fact]
        public void SelectionCanAlignAndDistributeShapes() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Align");
            VisioShape one = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "one", 2, 2, "One");
            VisioShape two = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "two", 8, 6, "Two");
            VisioShape three = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "three", 14, 4, "Three");
            one.Data["Row"] = "A";
            two.Data["Row"] = "A";
            three.Data["Row"] = "A";

            VisioShapeSelection selection = page.SelectWithData("Row", "A");
            selection.Align(VisioHorizontalAlignment.Left);

            double left = one.GetShapeBounds().Left;
            Assert.Equal(left, two.GetShapeBounds().Left, 6);
            Assert.Equal(left, three.GetShapeBounds().Left, 6);

            one.PinX = 2;
            two.PinX = 11;
            three.PinX = 14;
            selection.DistributeHorizontally();

            Assert.Equal(8, two.PinX, 6);

            selection.Align(VisioVerticalAlignment.Top);
            double top = one.GetShapeBounds().Top;
            Assert.Equal(top, two.GetShapeBounds().Top, 6);
            Assert.Equal(top, three.GetShapeBounds().Top, 6);
        }

        [Fact]
        public void SelectionCanRelayoutAsGridAndRerouteInternalConnectors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Relayout", 11, 8.5);
            VisioShape one = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "one", 7, 2, 2, 1, "One");
            VisioShape two = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"), "two", 2, 6, 1.4, 1.4, "Two");
            VisioShape three = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "three", 9, 6, 1.6, 1.0, "Three");
            VisioShape four = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "four", 4, 3, 2.2, 0.8, "Four");
            VisioShape outside = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "outside", 9, 1, "Outside");
            foreach (VisioShape shape in new[] { one, two, three, four }) {
                shape.Data["Group"] = "A";
            }

            VisioConnector first = page.AddConnector(one, two, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            VisioConnector second = page.AddConnector(two, three, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            VisioConnector external = page.AddConnector(three, outside, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);

            VisioShapeSelection selection = page.SelectWithData("Group", "A");
            selection.RelayoutAsGrid(new VisioSelectionLayoutOptions {
                Columns = 2,
                HorizontalSpacing = 0.4,
                VerticalSpacing = 0.3,
                Order = VisioSelectionLayoutOrder.TopLeftToBottomRight
            });

            Assert.Equal(2.4, two.PinX, 6);
            Assert.Equal(6.0, two.PinY, 6);
            Assert.Equal(4.9, three.PinX, 6);
            Assert.Equal(6.0, three.PinY, 6);
            Assert.Equal(2.4, four.PinX, 6);
            Assert.Equal(4.5, four.PinY, 6);
            Assert.Equal(4.9, one.PinX, 6);
            Assert.Equal(4.5, one.PinY, 6);
            Assert.Equal(9, outside.PinX, 6);
            Assert.Equal(1, outside.PinY, 6);
            Assert.Equal(2, first.Waypoints.Count);
            Assert.Equal(2, second.Waypoints.Count);
            Assert.Empty(external.Waypoints);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(4, loaded.Pages[0].ShapesWithData("Group", "A").Count);
            Assert.Single(loaded.Pages[0].Connectors, connector => connector.From.Id == "one" && connector.To.Id == "two");
        }

        [Fact]
        public void SelectionCanRelayoutAsVerticalStackWithoutReroutingConnectors() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Stack");
            VisioShape one = page.AddRectangle(2, 6, 1, 0.5, "One");
            VisioShape two = page.AddRectangle(5, 5, 2, 1, "Two");
            VisioShape three = page.AddRectangle(7, 3, 1.5, 0.75, "Three");
            VisioConnector connector = page.AddConnector(one, two, ConnectorKind.Dynamic);

            new VisioShapeSelection(new[] { one, two, three })
                .RelayoutAsVerticalStack(0.25, routeInternalConnectors: false);

            Assert.Equal(2.5, one.PinX, 6);
            Assert.Equal(6.0, one.PinY, 6);
            Assert.Equal(2.5, two.PinX, 6);
            Assert.Equal(5.0, two.PinY, 6);
            Assert.Equal(2.5, three.PinX, 6);
            Assert.Equal(3.875, three.PinY, 6);
            Assert.Empty(connector.Waypoints);
        }

        [Fact]
        public void ResizeToTextUsesDeterministicDrawingMeasurement() {
            VisioShape shortShape = new("short", 2, 2, 0.2, 0.2, string.Empty);
            shortShape.ResizeToText(minimumWidth: 0.75, minimumHeight: 0.4);

            Assert.Equal(0.75, shortShape.Width);
            Assert.True(shortShape.Height >= 0.4);
            Assert.Equal(shortShape.Width / 2, shortShape.LocPinX);
            Assert.Equal(shortShape.Height / 2, shortShape.LocPinY);

            VisioShape longShape = new("long", 2, 2, 0.2, 0.2, "Very long process label\nwith two lines");
            longShape.ResizeToText(new OfficeFontInfo("Calibri", 12), horizontalPadding: 0.2, verticalPadding: 0.1);

            Assert.True(longShape.Width > 2.0);
            Assert.True(longShape.Height > 0.55);
            Assert.Equal(longShape.Width / 2, longShape.LocPinX);
            Assert.Equal(longShape.Height / 2, longShape.LocPinY);
        }

        [Fact]
        public void ConnectorLabelCanResizeToTextAndRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("LabelSize");
            VisioShape source = page.AddRectangle(1, 4, 1.4, 0.7, "Source");
            VisioShape target = page.AddRectangle(5, 4, 1.4, 0.7, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            connector.Label = "Review, approve,\nor return";
            connector
                .PlaceLabel(0.55, offsetY: 0.2, width: 0.2, height: 0.1)
                .ResizeLabelToText(new OfficeFontInfo("Calibri", 12), maximumWidth: 1.4);

            Assert.NotNull(connector.LabelPlacement);
            Assert.True(connector.LabelPlacement!.Width > 0.2);
            Assert.True(connector.LabelPlacement.Width <= 1.4);
            Assert.True(connector.LabelPlacement.Height > 0.4);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.NotNull(loadedConnector.LabelPlacement);
            Assert.Equal(connector.LabelPlacement.Width, loadedConnector.LabelPlacement!.Width, 6);
            Assert.Equal(connector.LabelPlacement.Height, loadedConnector.LabelPlacement.Height, 6);
        }

        [Fact]
        public void ConnectorSelectionCanResizeLabelsToTextUsingConnectorTextStyle() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("SelectionLabels");
            VisioShape one = page.AddRectangle(1, 4, 1.4, 0.7, "One");
            VisioShape two = page.AddRectangle(3, 4, 1.4, 0.7, "Two");
            VisioShape three = page.AddRectangle(5, 4, 1.4, 0.7, "Three");
            VisioConnector first = page.AddConnector(one, two, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            VisioConnector second = page.AddConnector(two, three, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            first.Label = "Short";
            second.Label = "A much longer connector label";
            first.PlaceLabel();
            second.PlaceLabel();
            first.TextStyle = new VisioTextStyle { FontFamily = "Calibri", Size = 8 };
            second.TextStyle = new VisioTextStyle { FontFamily = "Calibri", Size = 14 };

            new VisioConnectorSelection(new[] { first, second })
                .ResizeLabelsToText(maximumWidth: 2.0);

            Assert.True(first.LabelPlacement!.Width >= 0.45);
            Assert.True(second.LabelPlacement!.Height > first.LabelPlacement.Height);
            Assert.True(second.LabelPlacement.Width <= 2.0);
        }

        [Fact]
        public void ResolveConnectorLabelOverlapsMovesLabelsAwayFromUnrelatedShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("LabelShapeCleanup", 7, 5);
            VisioShape source = page.AddRectangle(1, 2, 0.8, 0.5, "Source");
            VisioShape obstacle = page.AddRectangle(3, 2, 1, 1, "Obstacle");
            VisioShape target = page.AddRectangle(5, 2, 0.8, 0.5, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3, 2, width: 1.2, height: 0.4);
            connector.Label = "covers node";

            Assert.Contains(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false
            }), issue => issue.Kind == "ConnectorLabelOverlapsShape" && issue.ShapeId == obstacle.Id);

            page.ResolveConnectorLabelOverlaps();

            Assert.NotEqual(2, connector.LabelPlacement!.AbsolutePinY!.Value);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false
            }), issue => issue.Kind == "ConnectorLabelOverlapsShape" && issue.ShapeId == obstacle.Id);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void ResolveConnectorLabelOverlapsSeparatesOverlappingConnectorLabels() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("LabelLabelCleanup", 7, 5);
            VisioShape left = page.AddRectangle(1, 2, 0.8, 0.5, "Left");
            VisioShape middle = page.AddRectangle(3, 2, 0.8, 0.5, "Middle");
            VisioShape right = page.AddRectangle(5, 2, 0.8, 0.5, "Right");
            VisioConnector first = page.AddConnector(left, middle, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3, 3.4, width: 1.2, height: 0.4);
            first.Label = "first";
            VisioConnector second = page.AddConnector(middle, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3, 3.4, width: 1.2, height: 0.4);
            second.Label = "second";

            Assert.Contains(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false
            }), issue => issue.Kind == "ConnectorLabelOverlap");

            page.ResolveConnectorLabelOverlaps();

            Assert.Equal(3.4, first.LabelPlacement!.AbsolutePinY!.Value, 6);
            Assert.NotEqual(3.4, second.LabelPlacement!.AbsolutePinY!.Value);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false
            }), issue => issue.Kind == "ConnectorLabelOverlap");
        }

        [Fact]
        public void PolishDiagramSizesLabelsResolvesCollisionsAndFitsPage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Polish", 12, 8);
            VisioShape source = page.AddRectangle(5, 4, 0.8, 0.5, "Source");
            VisioShape obstacle = page.AddRectangle(7, 4, 1, 1, "Obstacle");
            VisioShape target = page.AddRectangle(9, 4, 0.8, 0.5, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(7, 4, width: 0.2, height: 0.1);
            connector.Label = "review and approve";

            page.PolishDiagram(new VisioDiagramPolishOptions {
                MaximumConnectorLabelWidth = 1.1,
                FitHorizontalMargin = 0.5,
                FitVerticalMargin = 0.25
            });

            Assert.True(connector.LabelPlacement!.Width > 0.2);
            Assert.True(connector.LabelPlacement.Width <= 1.1);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false
            }), issue => issue.Kind == "ConnectorLabelOverlapsShape" && issue.ShapeId == obstacle.Id);

            OfficeIMO.Visio.VisioShapeBounds bounds = page.GetContentBounds();
            Assert.Equal(0.5, bounds.Left, 6);
            Assert.Equal(0.25, bounds.Bottom, 6);
            Assert.True(page.Width < 12);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void PolishDiagramsSkipsBackgroundPagesByDefault() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage background = document.AddBackgroundPage("Background", 10, 8);
            VisioShape backgroundShape = background.AddRectangle(8, 6, 1, 1, "Background");
            VisioPage foreground = document.AddPage("Foreground", 10, 8);
            foreground.SetBackgroundPage(background);
            VisioShape source = foreground.AddRectangle(6, 4, 0.8, 0.5, "Source");
            VisioShape target = foreground.AddRectangle(8, 4, 0.8, 0.5, "Target");
            VisioConnector connector = foreground.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            connector.Label = "foreground";
            connector.PlaceLabel(0.5, width: 0.2, height: 0.1);

            document.PolishDiagrams();

            Assert.Equal(8, backgroundShape.PinX);
            Assert.NotEqual(6, source.PinX);
            Assert.True(connector.LabelPlacement!.Width > 0.2);
        }

        [Fact]
        public void CenterContentKeepsPageSizeAndMovesTopLevelShapes() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Center", 10, 8);
            VisioShape first = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "first", 1, 1, "First");
            VisioShape second = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "second", 3, 2, "Second");

            page.CenterContent();

            OfficeIMO.Visio.VisioShapeBounds bounds = page.GetContentBounds();
            Assert.Equal(5, bounds.CenterX, 6);
            Assert.Equal(4, bounds.CenterY, 6);
            Assert.Equal(10, page.Width);
            Assert.Equal(8, page.Height);
            Assert.Equal(new[] { "first", "second" }, page.Shapes.Select(shape => shape.Id).ToArray());
            Assert.NotEqual(1, first.PinX);
            Assert.NotEqual(2, second.PinY);
        }
    }
}
