using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
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
            internalConnector.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String);
            internalConnector.Data["Owner"] = "Platform";
            internalConnector.Data["Protocol"] = "HTTPS";
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
            Assert.Equal("Platform", connectorCopy.GetShapeDataValue("Owner"));
            Assert.Equal("HTTPS", connectorCopy.Data["Protocol"]);
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
            Assert.All(loaded.Pages[0].Connectors.Where(connector => connector.Label == "approved"),
                connector => Assert.Equal("HTTPS", connector.GetShapeDataValue("Protocol")));
            Assert.All(loaded.Pages[0].Connectors.Where(connector => connector.Label == "approved"),
                connector => Assert.Equal("Platform", connector.GetShapeDataValue("Owner")));
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
        public void DuplicateShapesOptionsCreateFriendlyUniqueIds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Friendly IDs", 11, 8.5);
            VisioShape api = new("api", 2, 5, 1.4, 0.8, "API") { NameU = "Rectangle" };
            VisioShape database = new("db", 5, 5, 1.4, 0.8, "Database") { NameU = "Rectangle" };
            VisioShape existingCopy = new("api-copy", 2, 3, 1.4, 0.8, "Existing copy") { NameU = "Rectangle" };
            page.Shapes.Add(api);
            page.Shapes.Add(database);
            page.Shapes.Add(existingCopy);
            VisioConnector connector = page.AddConnector("route", api, database, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            connector.Label = "SQL";

            VisioShapeSelection duplicates = page.DuplicateShapes(new[] { api, database }, new VisioShapeDuplicationOptions {
                IdSuffix = "-copy",
                ConnectorIdSuffix = "-copy",
                OffsetX = 1.25,
                OffsetY = -0.5
            });

            Assert.Equal(2, duplicates.Count);
            VisioShape apiCopy = page.FindShapeById("api-copy-2")!;
            VisioShape databaseCopy = page.FindShapeById("db-copy")!;
            Assert.NotNull(apiCopy);
            Assert.NotNull(databaseCopy);
            Assert.Equal(api.PinX + 1.25, apiCopy.PinX, 6);
            Assert.Equal(api.PinY - 0.5, apiCopy.PinY, 6);

            VisioConnector connectorCopy = page.Connectors.Single(current => current.Id == "route-copy");
            Assert.Same(apiCopy, connectorCopy.From);
            Assert.Same(databaseCopy, connectorCopy.To);
            Assert.Equal("SQL", connectorCopy.Label);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.NotNull(loaded.Pages[0].FindShapeById("api-copy-2"));
            Assert.NotNull(loaded.Pages[0].FindShapeById("db-copy"));
            Assert.Contains(loaded.Pages[0].Connectors, current => current.Label == "SQL" && current.From.Id == "api-copy-2" && current.To.Id == "db-copy");
        }

        [Fact]
        public void DuplicateShapesCopiesTargetedCommentsToClonedTargets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Annotated copies", 11, 8.5);
            VisioShape api = new("api", 2, 5, 1.4, 0.8, "API");
            VisioShape database = new("db", 5, 5, 1.4, 0.8, "Database");
            page.Shapes.Add(api);
            page.Shapes.Add(database);
            VisioConnector connector = page.AddConnector("route", api, database, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            connector.Label = "SQL";
            page.AddComment(api, "Review API", "Operations", "OP");
            page.AddCommentToShape(connector.Id, "Review route", "Operations", "OP");
            page.AddComment("Page note", "Operations", "OP");

            page.DuplicateShapes(new[] { api, database }, new VisioShapeDuplicationOptions {
                IdSuffix = "-copy",
                ConnectorIdSuffix = "-copy"
            });

            Assert.Equal(5, page.Comments.Count);
            Assert.Contains(page.Comments, comment => comment.ShapeId == "api-copy" && comment.Text == "Review API");
            Assert.Contains(page.Comments, comment => comment.ShapeId == "route-copy" && comment.Text == "Review route");
            Assert.Single(page.Comments, comment => comment.ShapeId == null && comment.Text == "Page note");
            Assert.Equal(page.Comments.Count, page.Comments.Select(comment => comment.Id).Distinct().Count());

            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = Assert.Single(loaded.Pages);
            Assert.Contains(loadedPage.Comments, comment => comment.ShapeId == "api-copy" && comment.Text == "Review API");
            Assert.Contains(loadedPage.Comments, comment => comment.ShapeId == "route-copy" && comment.Text == "Review route");
        }

        [Fact]
        public void FluentDuplicateShapesKeepsCopiesAddressableForChaining() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document.AsFluent()
                .Page("Fluent copies", page => page
                    .Rect("api", 2, 5, 1.4, 0.8, "API")
                    .Rect("db", 5, 5, 1.4, 0.8, "Database")
                    .Connect("api", "db", VisioSide.Right, VisioSide.Left, connector => connector.Label("SQL"))
                    .DuplicateShapes(new[] { "api", "db" }, duplicates => duplicates.ShapeData("Copied", "Yes"))
                    .Shape("api-copy", shape => shape.Text("API Copy"))
                    .Shape("db-copy", shape => shape.Text("Database Copy"))
                    .Connect("api-copy", "db-copy", VisioSide.Right, VisioSide.Left, connector => connector.Label("copied route")))
                .End()
                .Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Yes", page.FindShapeById("api-copy")!.GetShapeDataValue("Copied"));
            Assert.Equal("API Copy", page.FindShapeById("api-copy")!.Text);
            Assert.Equal("Database Copy", page.FindShapeById("db-copy")!.Text);
            Assert.Contains(page.Connectors, connector => connector.Label == "SQL" && connector.From.Id == "api-copy" && connector.To.Id == "db-copy");
            Assert.Contains(page.Connectors, connector => connector.Label == "copied route" && connector.From.Id == "api-copy" && connector.To.Id == "db-copy");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = Assert.Single(loaded.Pages);
            Assert.Equal("Yes", loadedPage.FindShapeById("api-copy")!.GetShapeDataValue("Copied"));
            Assert.Contains(loadedPage.Connectors, connector => connector.Label == "copied route" && connector.From.Id == "api-copy" && connector.To.Id == "db-copy");
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
            page.AddComment("Page review complete", "Operations", "OP");
            page.AddComment(source, "Source review", "Operations", "OP");
            page.AddCommentToShape(connector.Id, "Route review", "Operations", "OP");

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
            Assert.Equal(3, duplicate.Comments.Count);
            Assert.Contains(duplicate.Comments, comment => comment.ShapeId == null && comment.Text == "Page review complete");
            Assert.Contains(duplicate.Comments, comment => comment.ShapeId == sourceCopy.Id && comment.Text == "Source review");
            Assert.Contains(duplicate.Comments, comment => comment.ShapeId == connectorCopy.Id && comment.Text == "Route review");

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedDuplicate = loaded.Pages.Single(p => p.Name == "Original Copy");
            Assert.Equal(2, loadedDuplicate.Shapes.Count);
            Assert.Single(loadedDuplicate.Connectors);
            Assert.Equal(2, loadedDuplicate.Connectors.Single().Waypoints.Count);
            Assert.Equal(VisioPageRouteStyle.FlowchartLeftToRight, loadedDuplicate.ConnectorRouteStyle);
            Assert.NotNull(loadedDuplicate.BackgroundPage);
            Assert.Equal(3, loadedDuplicate.Comments.Count);
            Assert.Contains(loadedDuplicate.Comments, comment => comment.ShapeId == loadedDuplicate.Shapes.Single(shape => shape.Text == "Source").Id && comment.Text == "Source review");
            Assert.Contains(loadedDuplicate.Comments, comment => comment.ShapeId == loadedDuplicate.Connectors.Single().Id && comment.Text == "Route review");
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

        [Fact]
        public void DuplicatePagePreservesBackgroundPageChainsWhenCopyingBackground() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage watermark = document.AddBackgroundPage("Watermark", 11, 8.5);
            watermark.AddRectangle(5.5, 4.25, 4, 0.2, "Base watermark");
            VisioPage brand = document.AddBackgroundPage("Brand", 11, 8.5);
            brand.SetBackgroundPage(watermark);
            brand.AddRectangle(5.5, 8, 9, 0.35, "Brand header");
            VisioPage page = document.AddPage("Architecture", 11, 8.5);
            page.SetBackgroundPage(brand);

            VisioPage duplicate = page.Duplicate(new VisioPageDuplicationOptions {
                Name = "Architecture copy",
                DuplicateBackgroundPage = true,
                BackgroundPageName = "Brand copy"
            });

            Assert.Same(duplicate.BackgroundPage, document.Pages.Single(current => current.Name == "Brand copy"));
            Assert.Same(watermark, duplicate.BackgroundPage!.BackgroundPage);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedWatermark = loaded.Pages.Single(current => current.Name == "Watermark");
            VisioPage loadedBrandCopy = loaded.Pages.Single(current => current.Name == "Brand copy");
            VisioPage loadedDuplicate = loaded.Pages.Single(current => current.Name == "Architecture copy");
            Assert.Same(loadedBrandCopy, loadedDuplicate.BackgroundPage);
            Assert.Same(loadedWatermark, loadedBrandCopy.BackgroundPage);
        }

        [Fact]
        public void DuplicatePageKeepsSparseConnectorSpacingAndLayoutGridSizingUnset() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string duplicatePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Sparse Settings", 11, 8.5);
            page.SetConnectorSpacing(0.2, 0.25, 0.3, 0.35);
            page.SetLayoutGridSizing(1.1, 1.2, 0.4, 0.45);
            document.Save();

            RewritePageSheetCells(filePath, "Sparse Settings", (pageSheet, ns) => {
                RemovePageSheetCells(pageSheet, ns, "LineToLineY", "LineToNodeX", "LineToNodeY", "BlockSizeY", "AvenueSizeX", "AvenueSizeY");
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages.Single(current => current.Name == "Sparse Settings");
            Assert.Equal(0.2, loadedPage.LineToLineX.GetValueOrDefault(), 6);
            Assert.Null(loadedPage.LineToLineY);
            Assert.Equal(1.1, loadedPage.LayoutBlockSizeX.GetValueOrDefault(), 6);
            Assert.Null(loadedPage.LayoutBlockSizeY);

            loadedPage.Duplicate("Sparse Settings Copy");
            loaded.Save(duplicatePath);

            AssertPageSheetCellState(duplicatePath, "Sparse Settings Copy", "LineToLineX", "0.2");
            AssertPageSheetCellState(duplicatePath, "Sparse Settings Copy", "LineToLineY", null);
            AssertPageSheetCellState(duplicatePath, "Sparse Settings Copy", "LineToNodeX", null);
            AssertPageSheetCellState(duplicatePath, "Sparse Settings Copy", "LineToNodeY", null);
            AssertPageSheetCellState(duplicatePath, "Sparse Settings Copy", "BlockSizeX", "1.1");
            AssertPageSheetCellState(duplicatePath, "Sparse Settings Copy", "BlockSizeY", null);
            AssertPageSheetCellState(duplicatePath, "Sparse Settings Copy", "AvenueSizeX", null);
            AssertPageSheetCellState(duplicatePath, "Sparse Settings Copy", "AvenueSizeY", null);
        }

        private static void RewritePageSheetCells(string filePath, string pageName, Action<XElement, XNamespace> mutatePageSheet) {
            using ZipArchive archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry pagesEntry = archive.GetEntry("visio/pages/pages.xml") ?? throw new InvalidOperationException("Missing pages.xml");
            XDocument pages;
            using (Stream stream = pagesEntry.Open()) {
                pages = XDocument.Load(stream);
            }

            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement page = pages.Root!.Elements(ns + "Page")
                .Single(element => (string?)element.Attribute("Name") == pageName);
            mutatePageSheet(page.Element(ns + "PageSheet")!, ns);

            pagesEntry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("visio/pages/pages.xml");
            using Stream replacementStream = replacement.Open();
            pages.Save(replacementStream);
        }

        private static void RemovePageSheetCells(XElement pageSheet, XNamespace ns, params string[] names) {
            foreach (string name in names) {
                pageSheet.Elements(ns + "Cell")
                    .Where(current => (string?)current.Attribute("N") == name)
                    .Remove();
            }
        }

        private static void AssertPageSheetCellState(string filePath, string pageName, string cellName, string? expectedValue) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pages = ReadXml(archive, "visio/pages/pages.xml");
            XElement pageSheet = pages.Root!.Elements(ns + "Page")
                .Single(element => (string?)element.Attribute("Name") == pageName)
                .Element(ns + "PageSheet")!;
            XElement? cell = pageSheet.Elements(ns + "Cell")
                .SingleOrDefault(current => (string?)current.Attribute("N") == cellName);
            if (expectedValue == null) {
                Assert.Null(cell);
            } else {
                Assert.NotNull(cell);
                Assert.Equal(expectedValue, cell!.Attribute("V")!.Value);
            }
        }

        private static XDocument ReadXml(ZipArchive archive, string entryName) {
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}
