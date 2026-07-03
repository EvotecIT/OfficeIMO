using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioFluentDocumentTests {
        [Fact]
        public void CanBuildDocumentFluently() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            VisioDocument result = document.AsFluent()
                .Page("Page1", 8.5, 11, VisioMeasurementUnit.Inches, p => { })
                .End();

            Assert.Same(document, result);
            Assert.Single(document.Pages);
            Assert.Equal("Page1", document.Pages[0].Name);
            document.Save();
        }

        [Fact]
        public void FluentShapesUsePageUnitsAndPersistNumericShapeIds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document.AsFluent()
                .Page("Metric", 10, 10, VisioMeasurementUnit.Centimeters, p => p
                    .Rect("box", 2.54, 2.54, 2.54, 2.54, "Box"))
                .End();
            Assert.Equal("box", document.Pages[0].Shapes[0].Id);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape shape = Assert.Single(loaded.Pages[0].Shapes);
            Assert.Equal(2.54, shape.Width.FromInches(VisioMeasurementUnit.Centimeters), 5);

            using ZipArchive archive = ZipFile.OpenRead(filePath);
            using Stream pageStream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageXml = XDocument.Load(pageStream);
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            string? savedId = pageXml.Root!.Element(v + "Shapes")!.Element(v + "Shape")!.Attribute("ID")?.Value;
            Assert.True(int.TryParse(savedId, out _));
        }

        [Fact]
        public void FluentConnectCanTargetSidesAndStyleLines() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document.AsFluent()
                .Page("Page1", p => p
                    .Rect("left", 1, 1, 2, 1, "Left")
                    .Rect("right", 5, 1, 2, 1, "Right")
                    .Connect("left", "right", VisioSide.Right, VisioSide.Left, c => c
                        .RightAngle()
                        .LineColor(Color.DarkBlue)
                        .ArrowEnd(EndArrow.Triangle)))
                .End();
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector connector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.Equal(ConnectorKind.RightAngle, connector.Kind);
            Assert.NotNull(connector.FromConnectionPoint);
            Assert.NotNull(connector.ToConnectionPoint);
            Assert.Equal(Color.DarkBlue, connector.LineColor);
            Assert.Equal(EndArrow.Triangle, connector.EndArrow);
        }

        [Fact]
        public void FluentConnectUsesPageScopedConnectorIdsAcrossDocuments() {
            static VisioDocument CreateDocument(string filePath) {
                VisioDocument document = VisioDocument.Create(filePath);
                document.AsFluent()
                    .Page("Page1", p => p
                        .Rect("1", 1, 1, 2, 1, "Left")
                        .Rect("2", 5, 1, 2, 1, "Right")
                        .Connect("1", "2"))
                    .End();
                return document;
            }

            VisioDocument first = CreateDocument(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioDocument second = CreateDocument(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            Assert.Equal("3", Assert.Single(first.Pages[0].Connectors).Id);
            Assert.Equal("3", Assert.Single(second.Pages[0].Connectors).Id);
        }

        [Fact]
        public void FluentPageCanResolveShapeOverlaps() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document.AsFluent()
                .Page("Page1", p => p
                    .Rect("first", 2, 3, 1.2, 0.8, "First")
                    .Rect("second", 2.2, 3, 1.2, 0.8, "Second")
                    .Rect("third", 2.4, 3, 1.2, 0.8, "Third")
                    .ResolveShapeOverlaps(step: 0.25, maxAttempts: 12))
                .End();

            VisioPage page = document.Pages[0];
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckPageBounds = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabels = false
            }), issue => issue.Kind == "ShapeOverlap");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void FluentPageCanAddTitleAndTextBoxAdornments() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document.AsFluent()
                .Page("Overview", p => p
                    .Title("System Overview")
                    .TextBox("caption", 4.25, 9.8, 5.5, 0.35, "Generated by OfficeIMO", shape => shape
                        .TextColor(Color.DarkBlue)
                        .FontSize(11)))
                .End();

            VisioPage page = Assert.Single(document.Pages);
            VisioShape title = Assert.Single(page.Shapes, shape => shape.Id == "title");
            VisioShape caption = Assert.Single(page.Shapes, shape => shape.Id == "caption");
            Assert.Equal("Text Box", title.NameU);
            Assert.Equal("Text Box", caption.NameU);
            Assert.True(title.PinY > caption.PinY);
            Assert.Equal(0, title.LinePattern);
            Assert.Equal(0, caption.FillPattern);
            Assert.Equal(Color.DarkBlue, caption.TextStyle!.Color);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Contains(loaded.Pages[0].Shapes, shape => shape.Id == "title" && shape.Text == "System Overview");
            Assert.Contains(loaded.Pages[0].Shapes, shape => shape.Id == "caption" && shape.Text == "Generated by OfficeIMO");
        }

        [Fact]
        public void SideSelectionUsesNamedSidePointEvenWhenCustomPointsExist() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document.AsFluent()
                .Page("Page1", p => p
                    .Rect("left", 1, 1, 2, 2, "Left")
                    .Rect("right", 5, 1, 2, 2, "Right"))
                .End();

            VisioShape left = document.Pages[0].Shapes[0];
            VisioShape right = document.Pages[0].Shapes[1];
            left.ConnectionPoints.Add(new VisioConnectionPoint(0.25, 0.25, 0, 0));
            right.ConnectionPoints.Add(new VisioConnectionPoint(1.75, 1.75, 0, 0));

            VisioConnector connector = document.Pages[0].AddConnector(left, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            Assert.Equal(left.Width, connector.FromConnectionPoint!.X, 5);
            Assert.Equal(left.Height / 2, connector.FromConnectionPoint.Y, 5);
            Assert.Equal(0, connector.ToConnectionPoint!.X, 5);
            Assert.Equal(right.Height / 2, connector.ToConnectionPoint.Y, 5);
        }

        [Fact]
        public void FluentCanEditLoadedExistingPageWithoutAddingDuplicatePages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.AsFluent()
                .Page("Operations", p => p
                    .Rect("api", 2, 4, 1.4, 0.8, "Legacy API")
                    .Rect("db", 5, 4, 1.4, 0.8, "Database")
                    .Shape("api", shape => shape.ShapeData("Owner", "Ops"))
                    .Shape("db", shape => shape.ShapeData("Owner", "Data"))
                    .Connect("api", "db", VisioSide.Right, VisioSide.Left, connector => connector.Label("SQL")))
                .End()
                .Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            loaded.AsFluent()
                .ExistingPage("Operations", page => page
                    .ShapesWithData("Owner", "Ops", selection => selection
                        .Fill(Color.LightBlue)
                        .ShapeData("Reviewed", "Yes", "Reviewed", VisioShapeDataType.Boolean))
                    .ShapesContainingText("Legacy", selection => selection.Text(shape => shape.Text!.Replace("Legacy", "Production")))
                    .Rect("monitor", 7, 4, 1.4, 0.8, "Monitor")
                    .Connect("api", "monitor", VisioSide.Right, VisioSide.Left, connector => connector
                        .RightAngle()
                        .ArrowEnd(EndArrow.Triangle)
                        .Label("metrics"))
                    .Connectors(selection => selection
                        .LineColor(Color.DarkBlue)
                        .LabelPosition()))
                .End()
                .Save(updatedPath);

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("Operations", page.Name);
            Assert.Contains(page.Shapes, shape => shape.Id == "monitor" && shape.Text == "Monitor");

            VisioShape api = page.FindShapeById("api")!;
            Assert.NotNull(api);
            Assert.Equal("Production API", api.Text);
            Assert.Equal("Yes", api.GetShapeDataValue("Reviewed"));
            Assert.Equal(Color.LightBlue, api.FillColor);
            Assert.Equal(2, page.Connectors.Count);
            Assert.All(page.Connectors, connector => Assert.Equal(Color.DarkBlue, connector.LineColor));
            Assert.Contains(page.Connectors, connector => connector.Label == "metrics" && connector.EndArrow == EndArrow.Triangle);

            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void FluentPageOrAddIsIdempotentForCreateOrEditWorkflows() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            document.AsFluent()
                .PageOrAdd("Inventory", page => page.Rect("server", 2, 4, 1, 1, "Server"))
                .PageOrAdd("Inventory", page => page.Shape("server", shape => shape.ShapeData("Owner", "Ops")))
                .End();

            VisioPage page = Assert.Single(document.Pages);
            VisioShape server = Assert.Single(page.Shapes);
            Assert.Equal("Inventory", page.Name);
            Assert.Equal("Ops", server.GetShapeDataValue("Owner"));
        }

        [Fact]
        public void FluentAdvancedSelectionsEditLoadedGeometryPathsAndConnectors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Network", page => page
                    .Rect("zone", 5, 4, 6, 4, "Runtime Zone")
                    .Rect("ingress", 2.5, 4, 1, 1, "Ingress")
                    .Rect("api", 4.5, 4, 1, 1, "API")
                    .Rect("db", 6.5, 4, 1, 1, "Database")
                    .Rect("outside", 9.5, 4, 1, 1, "Outside")
                    .Shape("api", shape => shape
                        .UserCell("Tier", "App")
                        .Hyperlink("https://example.org/api")
                        .Protect(protection => protection.Position()))
                    .Shape("db", shape => shape.Layer("Data"))
                    .Connect("ingress", "api", VisioSide.Right, VisioSide.Left, connector => connector
                        .RightAngle()
                        .Layer("traffic")
                        .Hyperlink("https://example.org/flow")
                        .Label("https"))
                    .Connect("api", "db", VisioSide.Right, VisioSide.Left, connector => connector
                        .RightAngle()
                        .Layer("traffic")
                        .Protect(protection => protection.Endpoints())
                        .Label("sql")))
                .End()
                .Save();

            VisioDocument.Load(filePath)
                .AsFluent()
                .ExistingPage("Network", page => page
                    .ShapesContainedIn("zone", selection => selection
                        .ShapeData("Zone", "Core")
                        .Fill(Color.LightGreen))
                    .PathBetween("ingress", "db", selection => selection
                        .ShapeData("CriticalPath", "Yes")
                        .Stroke(Color.Red, 0.03))
                    .ConnectedComponent("ingress", selection => selection.ShapeData("Component", "Runtime"))
                    .ShapesWithUserCell("Tier", "App", selection => selection.ShapeData("TierReviewed", "Yes"))
                    .ShapesWithHyperlink("https://example.org/api", selection => selection.Fill(Color.LightBlue))
                    .ShapesWithProtection(protection => protection.LockMoveX == true, selection => selection.ShapeData("Locked", "Position"))
                    .ConnectedConnectors("api", selection => selection
                        .LineColor(Color.DarkBlue)
                        .LabelPosition())
                    .OutgoingConnectors("api", selection => selection.EndArrow(EndArrow.Triangle))
                    .ConnectorsInLayer("traffic", selection => selection.LineWeight(0.04))
                    .ConnectorsWithHyperlink("https://example.org/flow", selection => selection.Label("observed"))
                    .ConnectorsWithProtection(protection => protection.LockBegin == true, selection => selection.LinePattern(2)))
                .End()
                .Save(updatedPath);

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            VisioShape ingress = page.FindShapeById("ingress")!;
            VisioShape api = page.FindShapeById("api")!;
            VisioShape database = page.FindShapeById("db")!;
            VisioShape outside = page.FindShapeById("outside")!;

            Assert.Equal("Core", ingress.GetShapeDataValue("Zone"));
            Assert.Equal("Core", database.GetShapeDataValue("Zone"));
            Assert.Null(outside.GetShapeDataValue("Zone"));
            Assert.Equal("Yes", api.GetShapeDataValue("CriticalPath"));
            Assert.Equal("Runtime", database.GetShapeDataValue("Component"));
            Assert.Null(outside.GetShapeDataValue("Component"));
            Assert.Equal("Yes", api.GetShapeDataValue("TierReviewed"));
            Assert.Equal("Position", api.GetShapeDataValue("Locked"));
            Assert.Equal(Color.LightBlue, api.FillColor);

            Assert.Equal(2, page.Connectors.Count);
            Assert.All(page.Connectors, connector => {
                Assert.Equal(Color.DarkBlue, connector.LineColor);
                Assert.Equal(0.04, connector.LineWeight);
                Assert.NotNull(connector.LabelPlacement);
            });

            VisioConnector ingressToApi = page.Connectors.Single(connector => connector.From.Id == "ingress" && connector.To.Id == "api");
            VisioConnector apiToDatabase = page.Connectors.Single(connector => connector.From.Id == "api" && connector.To.Id == "db");
            Assert.Equal("observed", ingressToApi.Label);
            Assert.Equal(EndArrow.Triangle, apiToDatabase.EndArrow);
            Assert.Equal(2, apiToDatabase.LinePattern);

            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void FluentReplaceMasterStandardizesLoadedStencilShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Standardize", page => page
                    .Stencil("source", VisioStencils.Flowchart.Get("process"), 2, 5, "Source")
                    .Stencil("task", VisioStencils.Flowchart.Get("process"), 5, 5, "Review")
                    .Stencil("archive", VisioStencils.Flowchart.Get("data"), 8, 5, "Archive")
                    .Shape("task", shape => shape
                        .Fill(Color.LightYellow)
                        .Stroke(Color.Orange, 0.02)
                        .Layer("Review")
                        .ShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String, "Owning team")
                        .UserCell("Stage", "Review", "STR")
                        .Hyperlink("https://example.org/review", "Review docs"))
                    .Connect("source", "task", VisioSide.Right, VisioSide.Left, connector => connector.Label("submit"))
                    .Connect("task", "archive", VisioSide.Right, VisioSide.Left, connector => connector.Label("store")))
                .End()
                .Save();

            VisioDocument.Load(filePath)
                .AsFluent()
                .ExistingPage("Standardize", page => page
                    .ReplaceMaster("task", VisioStencils.Flowchart.Get("decision"), resizeToMaster: true)
                    .ReplaceMastersByMaster("Process", VisioStencils.Flowchart.Get("preparation"), resizeToMaster: true))
                .End()
                .Save(updatedPath);

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            VisioShape source = page.FindShapeById("source")!;
            VisioShape task = page.FindShapeById("task")!;
            VisioShape archive = page.FindShapeById("archive")!;

            Assert.Equal("Preparation", source.MasterNameU);
            Assert.Equal("flow.preparation", source.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal(2.2, source.Width, 6);
            Assert.Equal(1.0, source.Height, 6);

            Assert.Equal("Decision", task.MasterNameU);
            Assert.Equal("flow.decision", task.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal(2.0, task.Width, 6);
            Assert.Equal(1.4, task.Height, 6);
            Assert.Equal(Color.LightYellow, task.FillColor);
            Assert.Equal(Color.Orange, task.LineColor);
            Assert.Contains("Review", task.LayerNames);
            Assert.Equal("Operations", task.GetShapeDataValue("Owner"));
            Assert.Equal("Owning team", task.FindShapeData("Owner")!.Prompt);
            Assert.Equal("Review", task.GetUserCellValue("Stage"));
            Assert.Equal("https://example.org/review", task.Hyperlinks.Single().Address);

            Assert.Equal("Data", archive.MasterNameU);
            Assert.Single(page.IncomingConnectors(task));
            Assert.Single(page.OutgoingConnectors(task));
            Assert.All(page.Connectors, connector => Assert.Contains(connector.Label, new[] { "submit", "store" }));

            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void FluentCanRelayoutLoadedContainerMembersAndPersistRoutedConnectors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Operations", page => page
                    .Rect("ingress", 7, 2, 1.2, 0.7, "Ingress")
                    .Rect("api", 3, 6, 1.4, 0.8, "API")
                    .Rect("worker", 8, 5, 1.4, 0.8, "Worker")
                    .Rect("audit", 6, 1, 1.2, 0.7, "Audit")
                    .Container("runtime", "Runtime", new[] { "ingress", "api", "worker" }, options => {
                        options.Margin = 0.2D;
                        options.HeadingHeight = 0.3D;
                    })
                    .Connect("ingress", "api", VisioSide.Top, VisioSide.Bottom, connector => connector.Label("https"))
                    .Connect("api", "worker", VisioSide.Right, VisioSide.Left, connector => connector.Label("queue"))
                    .Connect("worker", "audit", VisioSide.Bottom, VisioSide.Top, connector => connector.Label("events")))
                .End()
                .Save();

            VisioDocument.Load(filePath)
                .AsFluent()
                .ExistingPage("Operations", page => page
                    .RelayoutContainerMembers("runtime", options => {
                        options.Columns = 1;
                        options.HorizontalSpacing = 0D;
                        options.VerticalSpacing = 0.25D;
                        options.Order = VisioSelectionLayoutOrder.TopLeftToBottomRight;
                    })
                    .AlignShapes(VisioHorizontalAlignment.Center, new[] { "ingress", "api", "worker" }))
                .End()
                .Save(updatedPath);

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            VisioShape ingress = page.FindShapeById("ingress")!;
            VisioShape api = page.FindShapeById("api")!;
            VisioShape worker = page.FindShapeById("worker")!;
            VisioShape audit = page.FindShapeById("audit")!;
            VisioShape runtime = page.FindShapeById("runtime")!;

            Assert.Equal(api.PinX, ingress.PinX, 6);
            Assert.Equal(api.PinX, worker.PinX, 6);
            Assert.True(api.PinY > worker.PinY);
            Assert.True(worker.PinY > ingress.PinY);
            Assert.Equal(6, audit.PinX, 6);
            Assert.Equal(1, audit.PinY, 6);
            Assert.Equal(3, page.GetContainerMembers(runtime).Count);
            Assert.True(runtime.GetShapeBounds().Left < ingress.GetShapeBounds().Left);
            Assert.True(runtime.GetShapeBounds().Right > worker.GetShapeBounds().Right);
            Assert.NotEmpty(page.Connectors.Single(connector => connector.From.Id == "ingress" && connector.To.Id == "api").Waypoints);
            Assert.NotEmpty(page.Connectors.Single(connector => connector.From.Id == "api" && connector.To.Id == "worker").Waypoints);
            Assert.Empty(page.Connectors.Single(connector => connector.From.Id == "worker" && connector.To.Id == "audit").Waypoints);

            Assert.Empty(VisioValidator.Validate(updatedPath));
        }
    }
}
