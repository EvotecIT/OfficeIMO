using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioDuplicationTests {
        [Fact]
        public void DuplicateSelectionCopiesMetadataAndInternalConnectors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Duplicate", 11, 8.5);

            VisioShape source = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "source", 2, 6, "Source");
            source.FillColor = Color.LightBlue;
            source.LineColor = Color.DodgerBlue;
            source.LineWeight = 0.03;
            source.LayerNames.Add("Ops");
            source.AddHyperlink("https://example.org/source", "Source docs");
            source.SetUserCell("Environment", "Prod", "STR");
            source.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String, "Owning team");
            source.PlacementStyle = VisioPlacementStyle.CompactRightDown;
            source.AllowHorizontalConnectorRoutingThrough = false;
            source.Protection.Size().Text();

            VisioShape target = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"), "target", 5, 6, "Target?");
            target.Data["Owner"] = "Operations";
            VisioShape external = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "external", 8, 6, "External");

            VisioConnector internalConnector = page.AddConnector(source, target, ConnectorKind.RightAngle, VisioSide.Right, VisioSide.Left);
            internalConnector.Label = "approved";
            internalConnector.EndArrow = EndArrow.Triangle;
            internalConnector.LineColor = Color.DarkGreen;
            internalConnector.LineWeight = 0.025;
            internalConnector.LinePattern = 2;
            internalConnector.RouteStyle = VisioPageRouteStyle.FlowchartLeftToRight;
            internalConnector.RouteAppearance = VisioLineRouteExtension.Curved;
            internalConnector.LineJumpStyle = VisioLineJumpStyle.Square;
            internalConnector.LineJumpCode = VisioConnectorLineJumpCode.Always;
            internalConnector.HorizontalJumpDirection = VisioHorizontalLineJumpDirection.Up;
            internalConnector.VerticalJumpDirection = VisioVerticalLineJumpDirection.Right;
            internalConnector.RerouteBehavior = VisioConnectorRerouteBehavior.OnCrossover;
            internalConnector.Waypoints.Add(VisioConnectorWaypoint.At(3.2, 6.7));
            internalConnector.LabelPlacement = VisioConnectorLabelPlacement.At(3.5, 6.9);
            internalConnector.AddHyperlink("https://example.org/route", "Route docs");
            internalConnector.Protection.Endpoints();

            VisioConnector externalConnector = page.AddConnector(source, external, ConnectorKind.Dynamic, VisioSide.Bottom, VisioSide.Left);
            externalConnector.Label = "external";

            VisioShapeSelection duplicates = page.SelectWithData("Owner", "Operations").Duplicate(1.5, -0.75);

            Assert.Equal(2, duplicates.Count);
            Assert.Equal(5, page.Shapes.Count);
            Assert.Equal(3, page.Connectors.Count);

            VisioShape sourceCopy = duplicates.Single(shape => shape.Text == "Source");
            VisioShape targetCopy = duplicates.Single(shape => shape.Text == "Target?");
            VisioConnector connectorCopy = page.Connectors.Single(connector =>
                connector.Label == "approved" &&
                ReferenceEquals(connector.From, sourceCopy) &&
                ReferenceEquals(connector.To, targetCopy));

            Assert.NotEqual(source.Id, sourceCopy.Id);
            Assert.Equal(source.PinX + 1.5, sourceCopy.PinX, 6);
            Assert.Equal(source.PinY - 0.75, sourceCopy.PinY, 6);
            Assert.Equal(source.MasterNameU, sourceCopy.MasterNameU);
            Assert.Equal(Color.LightBlue, sourceCopy.FillColor);
            Assert.Equal(Color.DodgerBlue, sourceCopy.LineColor);
            Assert.Equal(0.03, sourceCopy.LineWeight);
            Assert.Contains("Ops", sourceCopy.LayerNames);
            Assert.Equal("https://example.org/source", sourceCopy.Hyperlinks.Single().Address);
            Assert.Equal("Prod", sourceCopy.GetUserCellValue("Environment"));
            Assert.Equal("Operations", sourceCopy.GetShapeDataValue("Owner"));
            Assert.Equal("Owning team", sourceCopy.FindShapeData("Owner")!.Prompt);
            Assert.Equal(VisioPlacementStyle.CompactRightDown, sourceCopy.PlacementStyle);
            Assert.False(sourceCopy.AllowHorizontalConnectorRoutingThrough);
            Assert.True(sourceCopy.Protection.LockWidth);
            Assert.True(sourceCopy.Protection.LockTextEdit);

            Assert.Equal(ConnectorKind.RightAngle, connectorCopy.Kind);
            Assert.Equal(EndArrow.Triangle, connectorCopy.EndArrow);
            Assert.Equal(Color.DarkGreen, connectorCopy.LineColor);
            Assert.Equal(0.025, connectorCopy.LineWeight);
            Assert.Equal(2, connectorCopy.LinePattern);
            Assert.Equal(VisioPageRouteStyle.FlowchartLeftToRight, connectorCopy.RouteStyle);
            Assert.Equal(VisioLineRouteExtension.Curved, connectorCopy.RouteAppearance);
            Assert.Equal(VisioLineJumpStyle.Square, connectorCopy.LineJumpStyle);
            Assert.Equal(VisioConnectorLineJumpCode.Always, connectorCopy.LineJumpCode);
            Assert.Equal(VisioHorizontalLineJumpDirection.Up, connectorCopy.HorizontalJumpDirection);
            Assert.Equal(VisioVerticalLineJumpDirection.Right, connectorCopy.VerticalJumpDirection);
            Assert.Equal(VisioConnectorRerouteBehavior.OnCrossover, connectorCopy.RerouteBehavior);
            Assert.Equal(4.7, connectorCopy.Waypoints.Single().X, 6);
            Assert.Equal(5.95, connectorCopy.Waypoints.Single().Y, 6);
            Assert.Equal(5.0, connectorCopy.LabelPlacement!.AbsolutePinX!.Value, 6);
            Assert.Equal(6.15, connectorCopy.LabelPlacement.AbsolutePinY!.Value, 6);
            Assert.Equal("https://example.org/route", connectorCopy.Hyperlinks.Single().Address);
            Assert.True(connectorCopy.Protection.LockBegin);
            Assert.True(connectorCopy.Protection.LockEnd);
            Assert.NotSame(internalConnector.FromConnectionPoint, connectorCopy.FromConnectionPoint);
            Assert.NotSame(internalConnector.ToConnectionPoint, connectorCopy.ToConnectionPoint);
            Assert.Single(page.Connectors, connector => connector.Label == "external");

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(4, loaded.Pages[0].ShapesWithShapeData("Owner", "Operations").Count);
            Assert.Equal(2, loaded.Pages[0].Connectors.Count(connector => connector.Label == "approved"));
            Assert.Single(loaded.Pages[0].Connectors, connector => connector.Label == "external");
        }

        [Fact]
        public void DuplicateShapesCanSkipInternalConnectors() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("No connectors");
            VisioShape first = page.AddRectangle(2, 4, 1, 0.6, "First");
            VisioShape second = page.AddRectangle(4, 4, 1, 0.6, "Second");
            page.AddConnector(first, second, ConnectorKind.Dynamic);

            VisioShapeSelection duplicates = page.DuplicateShapes(new[] { first, second }, includeInternalConnectors: false);

            Assert.Equal(2, duplicates.Count);
            Assert.Equal(4, page.Shapes.Count);
            Assert.Single(page.Connectors);
        }

        [Fact]
        public void DuplicatePageCopiesSettingsLayersShapesAndConnectors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage background = document.AddBackgroundPage("Watermark", 11, 8.5);
            background.AddRectangle(5.5, 4.25, 9, 0.2, "Confidential");

            VisioPage page = document.AddPage("Original", 11, 8.5);
            page.SetBackgroundPage(background);
            page.DefaultUnit = VisioMeasurementUnit.Centimeters;
            page.ScaleMeasurementUnit = VisioMeasurementUnit.Centimeters;
            page.ViewScale = 0.75;
            page.ViewCenterX = 5.5;
            page.ViewCenterY = 4.25;
            page.PageLockDuplicate = true;
            page.DrawingSizeType = VisioDrawingSizeType.FitToDrawingContents;
            page.AutoResizeDrawing = false;
            page.AllowShapeSplitting = false;
            page.UiVisibility = VisioPageUiVisibility.Hidden;
            page.PlacementStyle = VisioPlacementStyle.HierarchyLeftToRightMiddle;
            page.PlacementDepth = VisioPlacementDepth.Deep;
            page.PlacementFlip = VisioPlacementFlip.Horizontal;
            page.MoveShapesAwayOnDrop = true;
            page.ResizePageToFitLayout = true;
            page.EnableLayoutGrid = true;
            page.ConnectorRouteStyle = VisioPageRouteStyle.FlowchartLeftToRight;
            page.ConnectorRouteAppearance = VisioLineRouteExtension.Curved;
            page.LineJumpStyle = VisioLineJumpStyle.Square;
            page.LineJumpCode = VisioLineJumpCode.DisplayOrder;
            page.HorizontalLineJumpDirection = VisioHorizontalLineJumpDirection.Up;
            page.VerticalLineJumpDirection = VisioVerticalLineJumpDirection.Right;
            page.PrintOrientation = VisioPagePrintOrientation.Landscape;
            page.SetMargins(0.4, 0.5, 0.6, 0.7);
            page.SetConnectorSpacing(0.2, 0.25, 0.3, 0.35);
            page.SetLayoutGridSizing(1.1, 1.2, 0.4, 0.45);
            page.AddLayer("Ops").Color = 4;

            VisioShape source = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "source", 2, 5, "Source");
            source.LayerNames.Add("Ops");
            source.SetShapeData("Owner", "Ops");
            VisioShape target = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"), "target", 6, 5, "Ready?");
            target.LayerNames.Add("Ops");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteThrough(VisioConnectorWaypoint.At(4, 5.5), VisioConnectorWaypoint.At(4, 4.5));
            connector.Label = "route";
            connector.EndArrow = EndArrow.Triangle;
            connector.LayerNames.Add("Ops");

            VisioPage duplicate = page.Duplicate("Original Copy");

            Assert.Equal(3, document.Pages.Count);
            Assert.Equal("Original Copy", duplicate.Name);
            Assert.False(duplicate.IsBackground);
            Assert.Same(background, duplicate.BackgroundPage);
            Assert.Equal(page.Width, duplicate.Width, 6);
            Assert.Equal(page.Height, duplicate.Height, 6);
            Assert.Equal(VisioMeasurementUnit.Centimeters, duplicate.DefaultUnit);
            Assert.Equal(0.75, duplicate.ViewScale, 6);
            Assert.True(duplicate.PageLockDuplicate);
            Assert.Equal(VisioDrawingSizeType.FitToDrawingContents, duplicate.DrawingSizeType);
            Assert.False(duplicate.AutoResizeDrawing);
            Assert.False(duplicate.AllowShapeSplitting);
            Assert.Equal(VisioPageUiVisibility.Hidden, duplicate.UiVisibility);
            Assert.Equal(VisioPlacementStyle.HierarchyLeftToRightMiddle, duplicate.PlacementStyle);
            Assert.Equal(VisioPlacementDepth.Deep, duplicate.PlacementDepth);
            Assert.Equal(VisioPlacementFlip.Horizontal, duplicate.PlacementFlip);
            Assert.True(duplicate.MoveShapesAwayOnDrop);
            Assert.True(duplicate.ResizePageToFitLayout);
            Assert.True(duplicate.EnableLayoutGrid);
            Assert.Equal(VisioPageRouteStyle.FlowchartLeftToRight, duplicate.ConnectorRouteStyle);
            Assert.Equal(VisioLineRouteExtension.Curved, duplicate.ConnectorRouteAppearance);
            Assert.Equal(VisioLineJumpStyle.Square, duplicate.LineJumpStyle);
            Assert.Equal(VisioLineJumpCode.DisplayOrder, duplicate.LineJumpCode);
            Assert.Equal(VisioHorizontalLineJumpDirection.Up, duplicate.HorizontalLineJumpDirection);
            Assert.Equal(VisioVerticalLineJumpDirection.Right, duplicate.VerticalLineJumpDirection);
            Assert.Equal(VisioPagePrintOrientation.Landscape, duplicate.PrintOrientation);
            Assert.Equal(0.4, duplicate.LeftMargin, 6);
            Assert.Equal(0.25, duplicate.LineToLineY!.Value, 6);
            Assert.Equal(1.2, duplicate.LayoutBlockSizeY!.Value, 6);
            Assert.Equal(0.45, duplicate.LayoutAvenueSizeY!.Value, 6);
            Assert.Single(duplicate.Layers);
            Assert.Equal("Ops", duplicate.Layers[0].Name);
            Assert.Equal(4, duplicate.Layers[0].Color);

            VisioShape sourceCopy = duplicate.Shapes.Single(shape => shape.Text == "Source");
            VisioShape targetCopy = duplicate.Shapes.Single(shape => shape.Text == "Ready?");
            Assert.NotEqual(source.Id, sourceCopy.Id);
            Assert.Equal(source.PinX, sourceCopy.PinX, 6);
            Assert.Equal("Ops", sourceCopy.GetShapeDataValue("Owner"));
            Assert.Contains("Ops", sourceCopy.LayerNames);

            VisioConnector connectorCopy = duplicate.Connectors.Single();
            Assert.NotEqual(connector.Id, connectorCopy.Id);
            Assert.Same(sourceCopy, connectorCopy.From);
            Assert.Same(targetCopy, connectorCopy.To);
            Assert.Equal("route", connectorCopy.Label);
            Assert.Equal(EndArrow.Triangle, connectorCopy.EndArrow);
            Assert.Equal(2, connectorCopy.Waypoints.Count);
            Assert.Equal(4, connectorCopy.Waypoints[0].X, 6);
            Assert.Contains("Ops", connectorCopy.LayerNames);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedDuplicate = loaded.Pages.Single(p => p.Name == "Original Copy");
            Assert.Equal(2, loadedDuplicate.Shapes.Count);
            Assert.Single(loadedDuplicate.Connectors);
            Assert.Equal(2, loadedDuplicate.Connectors.Single().Waypoints.Count);
            Assert.Equal(VisioPageRouteStyle.FlowchartLeftToRight, loadedDuplicate.ConnectorRouteStyle);
            Assert.NotNull(loadedDuplicate.BackgroundPage);
        }

        [Fact]
        public void DuplicatePageCanCopyBackgroundPageDependency() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage background = document.AddBackgroundPage("Brand", 11, 8.5);
            VisioShape header = background.AddRectangle(5.5, 8, 9, 0.35, "Brand header");
            header.FillColor = Color.LightGray;

            VisioPage page = document.AddPage("Architecture", 11, 8.5);
            page.SetBackgroundPage(background);
            VisioShape api = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "api", 3, 5, "API");
            VisioShape worker = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "worker", 7, 5, "Worker");
            page.AddConnector(api, worker, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteThrough(VisioConnectorWaypoint.At(5, 5.4), VisioConnectorWaypoint.At(5, 4.6));

            VisioPage duplicate = page.Duplicate(new VisioPageDuplicationOptions {
                Name = "Architecture copy",
                DuplicateBackgroundPage = true,
                BackgroundPageName = "Brand copy"
            });

            Assert.Equal(4, document.Pages.Count);
            VisioPage backgroundCopy = document.Pages.Single(p => p.Name == "Brand copy");
            Assert.True(backgroundCopy.IsBackground);
            Assert.NotSame(background, backgroundCopy);
            Assert.Single(backgroundCopy.Shapes);
            Assert.Equal("Brand header", backgroundCopy.Shapes[0].Text);
            Assert.NotEqual(header.Id, backgroundCopy.Shapes[0].Id);
            Assert.Same(backgroundCopy, duplicate.BackgroundPage);
            Assert.NotSame(background, duplicate.BackgroundPage);
            Assert.Single(duplicate.Connectors);
            Assert.Equal(2, duplicate.Connectors.Single().Waypoints.Count);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedBackgroundCopy = loaded.Pages.Single(p => p.Name == "Brand copy");
            VisioPage loadedDuplicate = loaded.Pages.Single(p => p.Name == "Architecture copy");
            VisioPage loadedOriginal = loaded.Pages.Single(p => p.Name == "Architecture");

            Assert.True(loadedBackgroundCopy.IsBackground);
            Assert.Same(loadedBackgroundCopy, loadedDuplicate.BackgroundPage);
            Assert.NotSame(loadedOriginal.BackgroundPage, loadedDuplicate.BackgroundPage);
            Assert.Equal(2, loadedDuplicate.Connectors.Single().Waypoints.Count);
        }
    }
}
