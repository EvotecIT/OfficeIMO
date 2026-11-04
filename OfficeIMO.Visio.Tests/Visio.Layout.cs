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
            VisioPage page = document.AddPage("Fit", OfficeLength.Inches(20), OfficeLength.Inches(20));
            VisioShape first = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "first", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), "First");
            first.Width = OfficeLength.Inches(2);
            first.Height = OfficeLength.Inches(1);
            VisioShape second = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "second", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(9, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Second");
            second.Width = OfficeLength.Inches(1);
            second.Height = OfficeLength.Inches(2);

            OfficeIMO.Visio.VisioShapeBounds before = page.GetContentBounds();

            Assert.Equal(3, before.Left.ToInches());
            Assert.Equal(1.5, before.Bottom.ToInches());
            Assert.Equal(9, before.Right.ToInches());
            Assert.Equal(5.5, before.Top.ToInches());

            page.FitToContent(horizontalMargin: 0.5, verticalMargin: 0.25);

            OfficeIMO.Visio.VisioShapeBounds after = page.GetContentBounds();
            Assert.Equal(0.5, after.Left.ToInches(), 6);
            Assert.Equal(0.25, after.Bottom.ToInches(), 6);
            Assert.Equal(6.5, after.Right.ToInches(), 6);
            Assert.Equal(4.25, after.Top.ToInches(), 6);
            Assert.Equal(7.0, page.Width.ToInches(), 6);
            Assert.Equal(4.5, page.Height.ToInches(), 6);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void ShapeBoundsUseLocPinAndRotation() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Transformed", OfficeLength.Inches(10), OfficeLength.Inches(10));
            VisioShape shape = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Rotated");
            shape.LocPinX = OfficeLength.Inches(0);
            shape.LocPinY = OfficeLength.Inches(0);
            shape.Angle = Math.PI / 2D;

            OfficeIMO.Visio.VisioShapeBounds bounds = shape.GetShapeBounds();

            Assert.Equal(3, bounds.Left.ToInches(), 6);
            Assert.Equal(4, bounds.Bottom.ToInches(), 6);
            Assert.Equal(4, bounds.Right.ToInches(), 6);
            Assert.Equal(6, bounds.Top.ToInches(), 6);
        }

        [Fact]
        public void FitToContentIncludesConnectorWaypointsAndLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("FitConnectors", OfficeLength.Inches(20), OfficeLength.Inches(20));
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Source");
            VisioShape target = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteThrough(
                    VisioConnectorWaypoint.At(2, 6),
                    VisioConnectorWaypoint.At(11, 6))
                .PlaceLabelAt(12, 6.5, width: 2, height: 0.5);
            connector.Label = "External label";

            OfficeIMO.Visio.VisioShapeBounds shapeOnlyBounds = page.GetContentBounds(includeConnectors: false);
            OfficeIMO.Visio.VisioShapeBounds fullBounds = page.GetContentBounds();

            Assert.Equal(3.5, shapeOnlyBounds.Left.ToInches());
            Assert.Equal(8.5, shapeOnlyBounds.Right.ToInches());
            Assert.Equal(2, fullBounds.Left.ToInches());
            Assert.Equal(13, fullBounds.Right.ToInches());
            Assert.Equal(6.75, fullBounds.Top.ToInches());

            page.FitToContent(horizontalMargin: 0.5, verticalMargin: 0.25);

            OfficeIMO.Visio.VisioShapeBounds fittedBounds = page.GetContentBounds();
            Assert.Equal(0.5, fittedBounds.Left.ToInches(), 6);
            Assert.Equal(0.25, fittedBounds.Bottom.ToInches(), 6);
            Assert.Equal(12.0, page.Width.ToInches(), 6);
            Assert.Equal(3.75, page.Height.ToInches(), 6);
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
            VisioPage page = document.AddPage("CenterConnectors", OfficeLength.Inches(14), OfficeLength.Inches(10));
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Source");
            VisioShape target = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteThrough(VisioConnectorWaypoint.At(3, 5))
                .PlaceLabelAt(3, 5.4, width: 1.4, height: 0.4);
            connector.Label = "Centered";

            page.CenterContent();

            OfficeIMO.Visio.VisioShapeBounds bounds = page.GetContentBounds();
            Assert.Equal(7, bounds.CenterX.ToInches(), 6);
            Assert.Equal(5, bounds.CenterY.ToInches(), 6);
            Assert.Equal(7, connector.Waypoints[0].X, 6);
            Assert.Equal(6.45, connector.Waypoints[0].Y, 6);
            Assert.Equal(7, connector.LabelPlacement!.AbsolutePinX!.Value, 6);
            Assert.Equal(6.85, connector.LabelPlacement.AbsolutePinY!.Value, 6);
        }

        [Fact]
        public void SelectionCanAlignAndDistributeShapes() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Align");
            VisioShape one = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "one", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "One");
            VisioShape two = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "two", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), "Two");
            VisioShape three = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "three", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(14, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), "Three");
            one.Data["Row"] = "A";
            two.Data["Row"] = "A";
            three.Data["Row"] = "A";

            VisioShapeSelection selection = page.SelectWithData("Row", "A");
            selection.Align(VisioHorizontalAlignment.Left);

            double left = one.GetShapeBounds().Left.ToInches();
            Assert.Equal(left, two.GetShapeBounds().Left.ToInches(), 6);
            Assert.Equal(left, three.GetShapeBounds().Left.ToInches(), 6);

            one.PinX = OfficeLength.Inches(2);
            two.PinX = OfficeLength.Inches(11);
            three.PinX = OfficeLength.Inches(14);
            selection.DistributeHorizontally();

            Assert.Equal(8, two.PinX.ToInches(), 6);

            selection.Align(VisioVerticalAlignment.Top);
            double top = one.GetShapeBounds().Top.ToInches();
            Assert.Equal(top, two.GetShapeBounds().Top.ToInches(), 6);
            Assert.Equal(top, three.GetShapeBounds().Top.ToInches(), 6);
        }

        [Fact]
        public void SelectionCanRelayoutAsGridAndRerouteInternalConnectors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Relayout", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            VisioShape one = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "one", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "One");
            VisioShape two = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"), "two", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), "Two");
            VisioShape three = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "three", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(9, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.6, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.0, page.DefaultUnit)), "Three");
            VisioShape four = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "four", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), "Four");
            VisioShape outside = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "outside", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(9, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Outside");
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

            Assert.Equal(2.4, two.PinX.ToInches(), 6);
            Assert.Equal(6.0, two.PinY.ToInches(), 6);
            Assert.Equal(4.9, three.PinX.ToInches(), 6);
            Assert.Equal(6.0, three.PinY.ToInches(), 6);
            Assert.Equal(2.4, four.PinX.ToInches(), 6);
            Assert.Equal(4.5, four.PinY.ToInches(), 6);
            Assert.Equal(4.9, one.PinX.ToInches(), 6);
            Assert.Equal(4.5, one.PinY.ToInches(), 6);
            Assert.Equal(9, outside.PinX.ToInches(), 6);
            Assert.Equal(1, outside.PinY.ToInches(), 6);
            Assert.Equal(2, first.Waypoints.Count);
            Assert.Equal(2, second.Waypoints.Count);
            Assert.Empty(external.Waypoints);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(4, loaded.Pages[0].ShapesWithData("Group", "A").Count);
            Assert.Single(loaded.Pages[0].Connectors, connector => connector.From?.Id == "one" && connector.To?.Id == "two");
        }

        [Fact]
        public void SelectionCanRelayoutAsVerticalStackWithoutReroutingConnectors() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Stack");
            VisioShape one = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "One");
            VisioShape two = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Two");
            VisioShape three = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.75, page.DefaultUnit)), "Three");
            VisioConnector connector = page.AddConnector(one, two, ConnectorKind.Dynamic);

            new VisioShapeSelection(new[] { one, two, three })
                .RelayoutAsVerticalStack(0.25, routeInternalConnectors: false);

            Assert.Equal(2.5, one.PinX.ToInches(), 6);
            Assert.Equal(6.0, one.PinY.ToInches(), 6);
            Assert.Equal(2.5, two.PinX.ToInches(), 6);
            Assert.Equal(5.0, two.PinY.ToInches(), 6);
            Assert.Equal(2.5, three.PinX.ToInches(), 6);
            Assert.Equal(3.875, three.PinY.ToInches(), 6);
            Assert.Empty(connector.Waypoints);
        }

        [Fact]
        public void SelectionRelayoutRejectsNonFiniteSpacing() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("InvalidSpacing");
            VisioShape one = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "One");
            VisioShape two = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Two");
            VisioShapeSelection selection = new(new[] { one, two }, page);

            ArgumentOutOfRangeException horizontal = Assert.Throws<ArgumentOutOfRangeException>(() =>
                selection.RelayoutAsGrid(new VisioSelectionLayoutOptions {
                    HorizontalSpacing = double.NaN
                }));
            ArgumentOutOfRangeException vertical = Assert.Throws<ArgumentOutOfRangeException>(() =>
                selection.RelayoutAsGrid(new VisioSelectionLayoutOptions {
                    VerticalSpacing = double.PositiveInfinity
                }));

            Assert.Contains("finite non-negative", horizontal.Message);
            Assert.Contains("finite non-negative", vertical.Message);
        }

        [Fact]
        public void ResizeToTextUsesDeterministicDrawingMeasurement() {
            VisioShape shortShape = new("short", OfficeLength.Inches(2), OfficeLength.Inches(2), OfficeLength.Inches(0.2), OfficeLength.Inches(0.2), string.Empty);
            shortShape.ResizeToText(minimumWidth: 0.75, minimumHeight: 0.4);

            Assert.Equal(0.75, shortShape.Width.ToInches());
            Assert.True(shortShape.Height.ToInches() >= 0.4);
            Assert.Equal(shortShape.Width.ToInches() / 2, shortShape.LocPinX.ToInches());
            Assert.Equal(shortShape.Height.ToInches() / 2, shortShape.LocPinY.ToInches());

            VisioShape longShape = new("long", OfficeLength.Inches(2), OfficeLength.Inches(2), OfficeLength.Inches(0.2), OfficeLength.Inches(0.2), "Very long process label\nwith two lines");
            longShape.ResizeToText(new OfficeFontInfo("Calibri", 12), horizontalPadding: 0.2, verticalPadding: 0.1);

            Assert.True(longShape.Width.ToInches() > 2.0);
            Assert.True(longShape.Height.ToInches() > 0.55);
            Assert.Equal(longShape.Width.ToInches() / 2, longShape.LocPinX.ToInches());
            Assert.Equal(longShape.Height.ToInches() / 2, longShape.LocPinY.ToInches());
        }

        [Fact]
        public void ConnectorLabelCanResizeToTextAndRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("LabelSize");
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.7, page.DefaultUnit)), "Source");
            VisioShape target = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.7, page.DefaultUnit)), "Target");
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
        public void ConnectorLabelResizeBreaksLongWordsWithinMaximumWidth() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("LongLabel");
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.7, page.DefaultUnit)), "Source");
            VisioShape target = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.7, page.DefaultUnit)), "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            connector.Label = "Supercalifragilisticexpialidocious";

            connector
                .PlaceLabel(0.5, width: 0.2, height: 0.1)
                .ResizeLabelToText(new OfficeFontInfo("Calibri", 12), maximumWidth: 0.8);

            Assert.NotNull(connector.LabelPlacement);
            Assert.True(connector.LabelPlacement!.Width <= 0.8);
            Assert.True(connector.LabelPlacement.Height > 0.7);
        }

        [Fact]
        public void ConnectorSelectionCanResizeLabelsToTextUsingConnectorTextStyle() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("SelectionLabels");
            VisioShape one = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.7, page.DefaultUnit)), "One");
            VisioShape two = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.7, page.DefaultUnit)), "Two");
            VisioShape three = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.7, page.DefaultUnit)), "Three");
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
            VisioPage page = document.AddPage("LabelShapeCleanup", OfficeLength.Inches(7), OfficeLength.Inches(5));
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Source");
            VisioShape obstacle = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Obstacle");
            VisioShape target = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Target");
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
            VisioPage page = document.AddPage("LabelLabelCleanup", OfficeLength.Inches(7), OfficeLength.Inches(5));
            VisioShape left = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Left");
            VisioShape middle = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Middle");
            VisioShape right = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Right");
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
        public void ResolveConnectorLabelOverlapsCanRunWholePageOptimizationPasses() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("DenseLabelCleanup", OfficeLength.Inches(8), OfficeLength.Inches(6));
            VisioShape left = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Left");
            VisioShape center = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Center");
            VisioShape right = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Right");
            VisioShape top = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Top");
            VisioShape bottom = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Bottom");
            VisioConnector first = page.AddConnector(left, center, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(4, 3.7, width: 1.2, height: 0.35);
            first.Label = "left-center";
            VisioConnector second = page.AddConnector(center, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(4, 3.7, width: 1.2, height: 0.35);
            second.Label = "center-right";
            VisioConnector third = page.AddConnector(top, bottom, ConnectorKind.Straight, VisioSide.Bottom, VisioSide.Top)
                .PlaceLabelAt(4, 3.7, width: 1.2, height: 0.35);
            third.Label = "top-bottom";

            Assert.Contains(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false
            }), issue => issue.Kind == "ConnectorLabelOverlap");

            page.ResolveConnectorLabelOverlaps(
                step: 0.2D,
                maxAttempts: 10,
                avoidShapes: false,
                avoidLabels: true,
                preferEndpointZones: false,
                avoidConnectorPaths: false,
                positionStep: 0.2D,
                maxPositionShifts: 2,
                optimizationPasses: 3);

            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false
            }), issue => issue.Kind == "ConnectorLabelOverlap");
            Assert.NotEqual(second.LabelPlacement!.AbsolutePinY, third.LabelPlacement!.AbsolutePinY);
        }

        [Fact]
        public void PolishDiagramValidatesConnectorLabelOptimizationPasses() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Invalid label polish");

            Assert.Throws<ArgumentOutOfRangeException>(() => page.PolishDiagram(new VisioDiagramPolishOptions {
                ConnectorLabelOptimizationPasses = 0
            }));
            Assert.Throws<ArgumentOutOfRangeException>(() => page.ResolveConnectorLabelOverlaps(optimizationPasses: 0));
        }

        [Fact]
        public void ResolveConnectorLabelOverlapsCanSlideLabelsAlongConnectorPaths() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("LabelPathCleanup", OfficeLength.Inches(8), OfficeLength.Inches(5));
            VisioShape left = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Left");
            VisioShape right = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Right");
            VisioConnector first = page.AddConnector(left, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabel(0.5, width: 1.0, height: 0.3);
            first.Label = "first";
            VisioConnector second = page.AddConnector(left, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabel(0.5, width: 1.0, height: 0.3);
            second.Label = "second";

            Assert.Contains(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false
            }), issue => issue.Kind == "ConnectorLabelOverlap");

            page.ResolveConnectorLabelOverlaps(
                step: 0.25D,
                maxAttempts: 0,
                avoidShapes: false,
                avoidLabels: true,
                preferEndpointZones: false,
                avoidConnectorPaths: false,
                positionStep: 0.25D,
                maxPositionShifts: 2);

            Assert.Equal(0.5D, first.LabelPlacement!.Position, 6);
            Assert.NotEqual(0.5D, second.LabelPlacement!.Position);
            Assert.Equal(0D, second.LabelPlacement.OffsetX, 6);
            Assert.Equal(0D, second.LabelPlacement.OffsetY, 6);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false
            }), issue => issue.Kind == "ConnectorLabelOverlap");
        }

        [Fact]
        public void PolishDiagramCanSlideConnectorLabelsAlongPaths() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("PolishLabelPathCleanup", OfficeLength.Inches(8), OfficeLength.Inches(5));
            VisioShape left = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Left");
            VisioShape right = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Right");
            VisioConnector first = page.AddConnector(left, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabel(0.5, width: 1.0, height: 0.3);
            first.Label = "first";
            VisioConnector second = page.AddConnector(left, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabel(0.5, width: 1.0, height: 0.3);
            second.Label = "second";

            page.PolishDiagram(new VisioDiagramPolishOptions {
                ConnectorLabelMaxAttempts = 0,
                ConnectorLabelPositionStep = 0.25D,
                ConnectorLabelMaxPositionShifts = 2,
                AvoidConnectorLabelShapeOverlaps = false,
                AvoidConnectorLabelConnectorPathOverlaps = false,
                FitToContent = false
            });

            Assert.NotEqual(0.5D, second.LabelPlacement!.Position);
            Assert.Equal(0D, second.LabelPlacement.OffsetY, 6);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false
            }), issue => issue.Kind == "ConnectorLabelOverlap");
        }

        [Fact]
        public void ResolveConnectorLabelOverlapsMovesLabelsAwayFromUnrelatedConnectorPaths() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("LabelConnectorCleanup", OfficeLength.Inches(7), OfficeLength.Inches(5));
            VisioShape left = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Left");
            VisioShape right = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Right");
            VisioShape top = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Top");
            VisioShape bottom = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Bottom");
            VisioConnector labeled = page.AddConnector(left, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3, 2, width: 1.2, height: 0.4);
            labeled.Label = "approval";
            VisioConnector crossing = page.AddConnector(top, bottom, ConnectorKind.Straight, VisioSide.Bottom, VisioSide.Top);

            Assert.True(ConnectorPathIntersectsBounds(crossing, GetConnectorLabelBounds(labeled)));
            Assert.Contains(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false,
                CheckConnectorLabelOverlaps = false
            }), issue =>
                issue.Kind == "ConnectorLabelCrossesConnector" &&
                issue.ConnectorId == labeled.Id &&
                issue.OtherConnectorId == crossing.Id);

            page.ResolveConnectorLabelOverlaps(step: 0.18D, maxAttempts: 8);

            Assert.NotEqual(3, labeled.LabelPlacement!.AbsolutePinX!.Value);
            Assert.False(ConnectorPathIntersectsBounds(crossing, GetConnectorLabelBounds(labeled)));
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false,
                CheckConnectorLabelOverlaps = false
            }), issue => issue.Kind == "ConnectorLabelCrossesConnector");
        }

        [Fact]
        public void ConnectorLabelPathAnalysisIgnoresUnroutedDynamicConnectors() {
            VisioDocument document = VisioDocument.Create(Path.Combine(
                Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("DynamicRoute", OfficeLength.Inches(7), OfficeLength.Inches(5));
            VisioShape left = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Left");
            VisioShape right = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Right");
            VisioShape top = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Top");
            VisioShape bottom = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Bottom");
            VisioConnector labeled = page.AddConnector(left, right,
                    ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3, 2, width: 1.2, height: 0.4);
            labeled.Label = "approval";
            VisioConnector dynamic = page.AddConnector(top, bottom,
                ConnectorKind.Dynamic, VisioSide.Bottom, VisioSide.Top);

            Assert.Empty(dynamic.Waypoints);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(
                new VisioDiagramQualityOptions {
                    CheckShapeOverlaps = false,
                    CheckConnectorShapeIntersections = false,
                    CheckConnectorLabelShapeOverlaps = false,
                    CheckConnectorLabelOverlaps = false
                }), issue => issue.Kind == "ConnectorLabelCrossesConnector"
                    && issue.OtherConnectorId == dynamic.Id);
        }

        [Fact]
        public void ConnectorLabelPathAnalysisRequiresStableLabeledConnectorPlacement() {
            VisioDocument document = VisioDocument.Create(Path.Combine(
                Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("DynamicLabel", OfficeLength.Inches(7), OfficeLength.Inches(5));
            VisioShape left = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Left");
            VisioShape right = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Right");
            VisioShape top = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Top");
            VisioShape bottom = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Bottom");
            VisioConnector dynamicLabel = page.AddConnector(left, right,
                    ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .PlaceLabel(0.5, width: 1.2, height: 0.4);
            dynamicLabel.Label = "approval";
            VisioConnector crossing = page.AddConnector(top, bottom,
                ConnectorKind.Straight, VisioSide.Bottom, VisioSide.Top);
            var options = new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false,
                CheckConnectorLabelOverlaps = false
            };

            Assert.DoesNotContain(page.AnalyzeVisualQuality(options), issue =>
                issue.Kind == "ConnectorLabelCrossesConnector"
                && issue.ConnectorId == dynamicLabel.Id
                && issue.OtherConnectorId == crossing.Id);

            dynamicLabel.PlaceLabelAt(3, 2, width: 1.2, height: 0.4);
            Assert.Contains(page.AnalyzeVisualQuality(options), issue =>
                issue.Kind == "ConnectorLabelCrossesConnector"
                && issue.ConnectorId == dynamicLabel.Id
                && issue.OtherConnectorId == crossing.Id);
        }

        [Fact]
        public void PolishDiagramMovesConnectorLabelsAwayFromConnectorPaths() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("PolishLabelConnectors", OfficeLength.Inches(7), OfficeLength.Inches(5));
            VisioShape left = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Left");
            VisioShape right = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Right");
            VisioShape top = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Top");
            VisioShape bottom = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Bottom");
            VisioConnector labeled = page.AddConnector(left, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3, 2, width: 1.2, height: 0.4);
            labeled.Label = "polished";
            VisioConnector crossing = page.AddConnector(top, bottom, ConnectorKind.Straight, VisioSide.Bottom, VisioSide.Top);

            page.PolishDiagram(new VisioDiagramPolishOptions {
                ConnectorLabelStep = 0.18D,
                ConnectorLabelMaxAttempts = 8,
                FitToContent = false
            });

            Assert.False(ConnectorPathIntersectsBounds(crossing, GetConnectorLabelBounds(labeled)));
        }

        [Fact]
        public void ResolveConnectorLabelOverlapsIgnoresGeneratedAdornmentShapes() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("LabelAdornmentCleanup", OfficeLength.Inches(7), OfficeLength.Inches(5));
            VisioShape adornment = page.AddTextBox("zone-caption", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.35, page.DefaultUnit)), "Zone caption");
            VisioSemanticUserCells.MarkGeneratedAdornment(adornment);
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Source");
            VisioShape target = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3, 2, width: 1.2, height: 0.35);
            connector.Label = "inside zone";

            Assert.True(adornment.IsDiagramAdornment);

            page.ResolveConnectorLabelOverlaps();

            Assert.Equal(3, connector.LabelPlacement!.AbsolutePinX!.Value, 6);
            Assert.Equal(2, connector.LabelPlacement.AbsolutePinY!.Value, 6);
        }

        [Fact]
        public void ResolveConnectorLabelOverlapsMovesLabelsAwayFromUserTextBoxes() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("LabelUserTextCleanup", OfficeLength.Inches(7), OfficeLength.Inches(5));
            VisioShape note = page.AddTextBox("user-caption", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.0, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.3, page.DefaultUnit)), "User note");
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Source");
            VisioShape target = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3, 2, width: 1.3, height: 0.4);
            connector.Label = "approval";

            Assert.False(VisioSemanticUserCells.IsGeneratedDiagramAdornment(note));
            Assert.Contains(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false
            }), issue => issue.Kind == "ConnectorLabelOverlapsShape" && issue.ShapeId == note.Id);

            page.ResolveConnectorLabelOverlaps();

            Assert.NotEqual(2, connector.LabelPlacement!.AbsolutePinY!.Value);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false
            }), issue => issue.Kind == "ConnectorLabelOverlapsShape" && issue.ShapeId == note.Id);
        }

        [Fact]
        public void ResolveConnectorLabelOverlapsCanPreferCommonEndpointZones() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("LabelZoneCleanup", OfficeLength.Inches(8), OfficeLength.Inches(5));
            VisioShape zone = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Shared zone");
            zone.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR");
            VisioShape unrelatedZone = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3.55, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), "Unrelated zone");
            unrelatedZone.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR");
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Source");
            VisioShape target = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabel(0.5, offsetY: 1.45, width: 1.2, height: 0.3);
            connector.Label = "zone label";

            Assert.False(Contains(zone.GetShapeBounds(), GetConnectorLabelBounds(connector)));

            page.ResolveConnectorLabelOverlaps(step: 0.25D, maxAttempts: 8, preferEndpointZones: true);

            Assert.True(Contains(zone.GetShapeBounds(), GetConnectorLabelBounds(connector)));
            Assert.Equal(0D, OverlapArea(GetConnectorLabelBounds(connector), unrelatedZone.GetShapeBounds()), 6);
        }

        [Fact]
        public void PolishDiagramCanPreferConnectorLabelsInsideEndpointZones() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("PolishLabelZones", OfficeLength.Inches(8), OfficeLength.Inches(5));
            VisioShape zone = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Shared zone");
            zone.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR");
            page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3.55, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), "Unrelated zone")
                .SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR");
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Source");
            VisioShape target = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabel(0.5, offsetY: 1.45, width: 1.2, height: 0.3);
            connector.Label = "polished zone label";

            page.PolishDiagram(new VisioDiagramPolishOptions {
                PreferConnectorLabelsInsideEndpointZones = true,
                ConnectorLabelStep = 0.25D,
                ConnectorLabelMaxAttempts = 8,
                FitToContent = false
            });

            Assert.True(Contains(zone.GetShapeBounds(), GetConnectorLabelBounds(connector)));
        }

        [Fact]
        public void ResolveShapeOverlapsMovesTopLevelShapesApart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("ShapeCleanup", OfficeLength.Inches(8), OfficeLength.Inches(6));
            VisioShape first = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), "First");
            VisioShape second = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.25, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), "Second");
            VisioShape third = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3.05, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), "Third");

            Assert.Contains(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckPageBounds = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabels = false
            }), issue => issue.Kind == "ShapeOverlap");

            page.ResolveShapeOverlaps(step: 0.25, maxAttempts: 12);

            Assert.Equal(2, first.PinX.ToInches());
            Assert.True(
                second.PinX.ToInches() != 2.25 ||
                second.PinY.ToInches() != 3 ||
                third.PinX.ToInches() != 2.1 ||
                third.PinY.ToInches() != 3.05);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckPageBounds = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabels = false
            }), issue => issue.Kind == "ShapeOverlap");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void PolishDiagramSizesLabelsResolvesCollisionsAndFitsPage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Polish", OfficeLength.Inches(12), OfficeLength.Inches(8));
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Source");
            VisioShape obstacle = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Obstacle");
            VisioShape target = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(9, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Target");
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
            Assert.Equal(0.5, bounds.Left.ToInches(), 6);
            Assert.Equal(0.25, bounds.Bottom.ToInches(), 6);
            Assert.True(page.Width.ToInches() < 12);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void PolishDiagramCanRouteConnectorsAroundShapesBeforeLabelCleanup() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Route polish", OfficeLength.Inches(8), OfficeLength.Inches(5));
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Source");
            VisioShape obstacle = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.0, page.DefaultUnit)), "Obstacle");
            VisioShape target = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, page.DefaultUnit)), "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabel(0.5, width: 1.0, height: 0.3);
            connector.Label = "approved";

            page.PolishDiagram(new VisioDiagramPolishOptions {
                ResolveConnectorShapeIntersections = true,
                FitToContent = false
            });

            Assert.Equal(ConnectorKind.RightAngle, connector.Kind);
            Assert.Equal(2, connector.Waypoints.Count);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorLabels = false
            }), issue => issue.Kind == "ConnectorCrossesShape" && issue.ShapeId == obstacle.Id);
        }

        [Fact]
        public void PolishDiagramCanResolveShapeOverlapsBeforeFittingPage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("PolishShapes", OfficeLength.Inches(12), OfficeLength.Inches(8));
            VisioShape first = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), "First");
            VisioShape second = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), "Second");
            VisioShape third = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, page.DefaultUnit)), "Third");

            page.PolishDiagram(new VisioDiagramPolishOptions {
                ResolveShapeOverlaps = true,
                ShapeOverlapStep = 0.25,
                ShapeOverlapMaxAttempts = 12,
                FitHorizontalMargin = 0.5,
                FitVerticalMargin = 0.25
            });

            Assert.Equal(0.5, page.GetContentBounds().Left.ToInches(), 6);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckPageBounds = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabels = false
            }), issue => issue.Kind == "ShapeOverlap");
            Assert.True(page.Width.ToInches() < 12);
            Assert.Equal(0.5 + first.Width.ToInches() / 2D, first.PinX.ToInches(), 6);
            Assert.True(second.PinX.ToInches() > first.PinX.ToInches() || third.PinX.ToInches() > first.PinX.ToInches());

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void PolishDiagramsSkipsBackgroundPagesByDefault() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage background = document.AddBackgroundPage("Background", OfficeLength.Inches(10), OfficeLength.Inches(8));
            VisioShape backgroundShape = background.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(8, background.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, background.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, background.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, background.DefaultUnit)), "Background");
            VisioPage foreground = document.AddPage("Foreground", OfficeLength.Inches(10), OfficeLength.Inches(8));
            foreground.SetBackgroundPage(background);
            VisioShape source = foreground.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, foreground.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, foreground.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, foreground.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, foreground.DefaultUnit)), "Source");
            VisioShape target = foreground.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(8, foreground.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, foreground.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.8, foreground.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.5, foreground.DefaultUnit)), "Target");
            VisioConnector connector = foreground.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            connector.Label = "foreground";
            connector.PlaceLabel(0.5, width: 0.2, height: 0.1);

            document.PolishDiagrams();

            Assert.Equal(8, backgroundShape.PinX.ToInches());
            Assert.NotEqual(6, source.PinX.ToInches());
            Assert.True(connector.LabelPlacement!.Width > 0.2);
        }

        [Fact]
        public void CenterContentKeepsPageSizeAndMovesTopLevelShapes() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Center", OfficeLength.Inches(10), OfficeLength.Inches(8));
            VisioShape first = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "first", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "First");
            VisioShape second = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "second", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Second");

            page.CenterContent();

            OfficeIMO.Visio.VisioShapeBounds bounds = page.GetContentBounds();
            Assert.Equal(5, bounds.CenterX.ToInches(), 6);
            Assert.Equal(4, bounds.CenterY.ToInches(), 6);
            Assert.Equal(10, page.Width.ToInches());
            Assert.Equal(8, page.Height.ToInches());
            Assert.Equal(new[] { "first", "second" }, page.Shapes.Select(shape => shape.Id).ToArray());
            Assert.NotEqual(1, first.PinX.ToInches());
            Assert.NotEqual(2, second.PinY.ToInches());
        }

        private static OfficeIMO.Visio.VisioShapeBounds GetConnectorLabelBounds(VisioConnector connector) {
            VisioConnectorLabelPlacement placement = connector.LabelPlacement ?? VisioConnectorLabelPlacement.Along(0.5D);
            double pinX;
            double pinY;
            if (placement.PinX.HasValue && placement.PinY.HasValue) {
                pinX = placement.PinX.Value;
                pinY = placement.PinY.Value;
            } else {
                (double startX, double startY) = ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint);
                (double endX, double endY) = ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint);
                double position = Math.Max(0D, Math.Min(1D, placement.Position));
                pinX = startX + ((endX - startX) * position) + placement.OffsetX;
                pinY = startY + ((endY - startY) * position) + placement.OffsetY;
            }

            double width = placement.Width;
            double height = placement.Height;
            return new OfficeIMO.Visio.VisioShapeBounds(
                OfficeLength.Inches(pinX - (width / 2D)),
                OfficeLength.Inches(pinY - (height / 2D)),
                OfficeLength.Inches(pinX + (width / 2D)),
                OfficeLength.Inches(pinY + (height / 2D)));
        }

        private static (double X, double Y) ResolveEndpoint(VisioShape shape, VisioShape other, VisioConnectionPoint? connectionPoint) {
            if (connectionPoint != null) {
                return shape.GetAbsolutePoint(connectionPoint.X, connectionPoint.Y);
            }

            OfficeIMO.Visio.VisioShapeBounds shapeBounds = shape.GetShapeBounds();
            OfficeIMO.Visio.VisioShapeBounds otherBounds = other.GetShapeBounds();
            double dx = otherBounds.CenterX.ToInches() - shapeBounds.CenterX.ToInches();
            double dy = otherBounds.CenterY.ToInches() - shapeBounds.CenterY.ToInches();
            return Math.Abs(dx) >= Math.Abs(dy)
                ? (dx >= 0 ? shapeBounds.Right.ToInches() : shapeBounds.Left.ToInches(), shapeBounds.CenterY.ToInches())
                : (shapeBounds.CenterX.ToInches(), dy >= 0 ? shapeBounds.Top.ToInches() : shapeBounds.Bottom.ToInches());
        }

        private static bool ConnectorPathIntersectsBounds(VisioConnector connector, OfficeIMO.Visio.VisioShapeBounds bounds) {
            (double startX, double startY) = ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint);
            (double endX, double endY) = ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint);
            if (connector.Waypoints.Count == 0) {
                return SegmentIntersectsBounds(startX, startY, endX, endY, bounds);
            }

            double currentX = startX;
            double currentY = startY;
            foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                if (SegmentIntersectsBounds(currentX, currentY, waypoint.X, waypoint.Y, bounds)) {
                    return true;
                }

                currentX = waypoint.X;
                currentY = waypoint.Y;
            }

            return SegmentIntersectsBounds(currentX, currentY, endX, endY, bounds);
        }

        private static bool SegmentIntersectsBounds(double startX, double startY, double endX, double endY, OfficeIMO.Visio.VisioShapeBounds bounds) {
            return OfficeGeometry.SegmentIntersectsRectangle(
                (startX, startY),
                (endX, endY),
                bounds.Left.ToInches(),
                bounds.Bottom.ToInches(),
                bounds.Right.ToInches(),
                bounds.Top.ToInches());
        }

        private static bool Contains(OfficeIMO.Visio.VisioShapeBounds outer, OfficeIMO.Visio.VisioShapeBounds inner) {
            const double tolerance = 1e-6;
            return outer.Left.ToInches() <= inner.Left.ToInches() + tolerance &&
                   outer.Bottom.ToInches() <= inner.Bottom.ToInches() + tolerance &&
                   outer.Right.ToInches() + tolerance >= inner.Right.ToInches() &&
                   outer.Top.ToInches() + tolerance >= inner.Top.ToInches();
        }

        private static double OverlapArea(OfficeIMO.Visio.VisioShapeBounds first, OfficeIMO.Visio.VisioShapeBounds second) {
            double width = Math.Max(0D, Math.Min(first.Right.ToInches(), second.Right.ToInches()) - Math.Max(first.Left.ToInches(), second.Left.ToInches()));
            double height = Math.Max(0D, Math.Min(first.Top.ToInches(), second.Top.ToInches()) - Math.Max(first.Bottom.ToInches(), second.Bottom.ToInches()));
            return width * height;
        }
    }
}
