using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.OfficeColor;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioFluentDocumentTests {
        [Fact]
        public void CanBuildDocumentFluently() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            VisioDocument result = document.AsFluent()
                .AddPage("Page1", OfficeLength.Inches(8.5), OfficeLength.Inches(11), VisioMeasurementUnit.Inches, configure: p => { })
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
                .AddPage("Metric", OfficeLength.Centimeters(10), OfficeLength.Centimeters(10), VisioMeasurementUnit.Centimeters, configure: p => p
                    .AddRectangle("box", OfficeLength.Centimeters(2.54), OfficeLength.Centimeters(2.54), OfficeLength.Centimeters(2.54), OfficeLength.Centimeters(2.54), "Box"))
                .End();
            Assert.Equal("box", document.Pages[0].Shapes[0].Id);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape shape = Assert.Single(loaded.Pages[0].Shapes);
            Assert.Equal(2.54, shape.Width.ToInches().FromInches(VisioMeasurementUnit.Centimeters), 5);

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
                .AddPage("Page1", p => p
                    .AddRectangle("left", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Left")
                    .AddRectangle("right", OfficeLength.Inches(5), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Right")
                    .AddConnector("left", "right", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: c => c
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
                    .AddPage("Page1", p => p
                        .AddRectangle("1", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Left")
                        .AddRectangle("2", OfficeLength.Inches(5), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Right")
                        .AddConnector("1", "2"))
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
                .AddPage("Page1", p => p
                    .AddRectangle("first", OfficeLength.Inches(2), OfficeLength.Inches(3), OfficeLength.Inches(1.2), OfficeLength.Inches(0.8), "First")
                    .AddRectangle("second", OfficeLength.Inches(2.2), OfficeLength.Inches(3), OfficeLength.Inches(1.2), OfficeLength.Inches(0.8), "Second")
                    .AddRectangle("third", OfficeLength.Inches(2.4), OfficeLength.Inches(3), OfficeLength.Inches(1.2), OfficeLength.Inches(0.8), "Third")
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
                .AddPage("Overview", p => p
                    .AddTitle("System Overview")
                    .AddTextBox("caption", OfficeLength.Inches(4.25), OfficeLength.Inches(9.8), OfficeLength.Inches(5.5), OfficeLength.Inches(0.35), "Generated by OfficeIMO", shape => shape
                        .TextColor(Color.DarkBlue)
                        .FontSize(11)))
                .End();

            VisioPage page = Assert.Single(document.Pages);
            VisioShape title = Assert.Single(page.Shapes, shape => shape.Id == "title");
            VisioShape caption = Assert.Single(page.Shapes, shape => shape.Id == "caption");
            Assert.Equal("Text Box", title.NameU);
            Assert.Equal("Text Box", caption.NameU);
            Assert.True(title.PinY.ToInches() > caption.PinY.ToInches());
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
                .AddPage("Page1", p => p
                    .AddRectangle("left", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(2), "Left")
                    .AddRectangle("right", OfficeLength.Inches(5), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(2), "Right"))
                .End();

            VisioShape left = document.Pages[0].Shapes[0];
            VisioShape right = document.Pages[0].Shapes[1];
            left.ConnectionPoints.Add(new VisioConnectionPoint(0.25, 0.25, 0, 0));
            right.ConnectionPoints.Add(new VisioConnectionPoint(1.75, 1.75, 0, 0));

            VisioConnector connector = document.Pages[0].AddConnector(left, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            Assert.Equal(left.Width.ToInches(), connector.FromConnectionPoint!.X, 5);
            Assert.Equal(left.Height.ToInches() / 2, connector.FromConnectionPoint.Y, 5);
            Assert.Equal(0, connector.ToConnectionPoint!.X, 5);
            Assert.Equal(right.Height.ToInches() / 2, connector.ToConnectionPoint.Y, 5);
        }

        [Fact]
        public void FluentCanEditLoadedExistingPageWithoutAddingDuplicatePages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.AsFluent()
                .AddPage("Operations", p => p
                    .AddRectangle("api", OfficeLength.Inches(2), OfficeLength.Inches(4), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "Legacy API")
                    .AddRectangle("db", OfficeLength.Inches(5), OfficeLength.Inches(4), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "Database")
                    .GetShape("api", shape => shape.ShapeData("Owner", "Ops"))
                    .GetShape("db", shape => shape.ShapeData("Owner", "Data"))
                    .AddConnector("api", "db", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => connector.Label("SQL")))
                .End()
                .Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            loaded.AsFluent()
                .GetPage("Operations", page => page
                    .ShapesWithData("Owner", "Ops", selection => selection
                        .Fill(Color.LightBlue)
                        .ShapeData("Reviewed", "Yes", "Reviewed", VisioShapeDataType.Boolean))
                    .ShapesContainingText("Legacy", selection => selection.Text(shape => shape.Text!.Replace("Legacy", "Production")))
                    .AddRectangle("monitor", OfficeLength.Inches(7), OfficeLength.Inches(4), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "Monitor")
                    .AddConnector("api", "monitor", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => connector
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
                .GetOrAddPage("Inventory", page => page.AddRectangle("server", OfficeLength.Inches(2), OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(1), "Server"))
                .GetOrAddPage("Inventory", page => page.GetShape("server", shape => shape.ShapeData("Owner", "Ops")))
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
                .AddPage("Network", page => page
                    .AddRectangle("zone", OfficeLength.Inches(5), OfficeLength.Inches(4), OfficeLength.Inches(6), OfficeLength.Inches(4), "Runtime Zone")
                    .AddRectangle("ingress", OfficeLength.Inches(2.5), OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(1), "Ingress")
                    .AddRectangle("api", OfficeLength.Inches(4.5), OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(1), "API")
                    .AddRectangle("db", OfficeLength.Inches(6.5), OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(1), "Database")
                    .AddRectangle("outside", OfficeLength.Inches(9.5), OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(1), "Outside")
                    .GetShape("api", shape => shape
                        .UserCell("Tier", "App")
                        .Hyperlink("https://example.org/api")
                        .Protect(protection => protection.Position()))
                    .GetShape("db", shape => shape.Layer("Data"))
                    .AddConnector("ingress", "api", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => connector
                        .RightAngle()
                        .Layer("traffic")
                        .Hyperlink("https://example.org/flow")
                        .Label("https"))
                    .AddConnector("api", "db", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => connector
                        .RightAngle()
                        .Layer("traffic")
                        .Protect(protection => protection.Endpoints())
                        .Label("sql")))
                .End()
                .Save();

            VisioDocument.Load(filePath)
                .AsFluent()
                .GetPage("Network", page => page
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

            VisioConnector ingressToApi = page.Connectors.Single(connector => connector.From?.Id == "ingress" && connector.To?.Id == "api");
            VisioConnector apiToDatabase = page.Connectors.Single(connector => connector.From?.Id == "api" && connector.To?.Id == "db");
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
                .AddPage("Standardize", page => page
                    .AddStencilShape(id: "source", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(2), y: OfficeLength.Inches(5), text: "Source")
                    .AddStencilShape(id: "task", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(5), y: OfficeLength.Inches(5), text: "Review")
                    .AddStencilShape(id: "archive", stencil: VisioStencils.Flowchart.Get("data"), x: OfficeLength.Inches(8), y: OfficeLength.Inches(5), text: "Archive")
                    .GetShape("task", shape => shape
                        .Fill(Color.LightYellow)
                        .Stroke(Color.Orange, 0.02)
                        .Layer("Review")
                        .ShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String, "Owning team")
                        .UserCell("Stage", "Review", "STR")
                        .Hyperlink("https://example.org/review", "Review docs"))
                    .AddConnector("source", "task", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => connector.Label("submit"))
                    .AddConnector("task", "archive", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => connector.Label("store")))
                .End()
                .Save();

            VisioDocument.Load(filePath)
                .AsFluent()
                .GetPage("Standardize", page => page
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
            Assert.Equal(2.2, source.Width.ToInches(), 6);
            Assert.Equal(1.0, source.Height.ToInches(), 6);

            Assert.Equal("Decision", task.MasterNameU);
            Assert.Equal("flow.decision", task.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal(2.0, task.Width.ToInches(), 6);
            Assert.Equal(1.4, task.Height.ToInches(), 6);
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
                .AddPage("Operations", page => page
                    .AddRectangle("ingress", OfficeLength.Inches(7), OfficeLength.Inches(2), OfficeLength.Inches(1.2), OfficeLength.Inches(0.7), "Ingress")
                    .AddRectangle("api", OfficeLength.Inches(3), OfficeLength.Inches(6), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "API")
                    .AddRectangle("worker", OfficeLength.Inches(8), OfficeLength.Inches(5), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "Worker")
                    .AddRectangle("audit", OfficeLength.Inches(6), OfficeLength.Inches(1), OfficeLength.Inches(1.2), OfficeLength.Inches(0.7), "Audit")
                    .AddContainer("runtime", "Runtime", new[] { "ingress", "api", "worker" }, options => {
                        options.Margin = 0.2D;
                        options.HeadingHeight = 0.3D;
                    })
                    .AddConnector("ingress", "api", fromSide: VisioSide.Top, toSide: VisioSide.Bottom, configure: connector => connector.Label("https"))
                    .AddConnector("api", "worker", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => connector.Label("queue"))
                    .AddConnector("worker", "audit", fromSide: VisioSide.Bottom, toSide: VisioSide.Top, configure: connector => connector.Label("events")))
                .End()
                .Save();

            VisioDocument.Load(filePath)
                .AsFluent()
                .GetPage("Operations", page => page
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

            Assert.Equal(api.PinX.ToInches(), ingress.PinX.ToInches(), 6);
            Assert.Equal(api.PinX.ToInches(), worker.PinX.ToInches(), 6);
            Assert.True(api.PinY.ToInches() > worker.PinY.ToInches());
            Assert.True(worker.PinY.ToInches() > ingress.PinY.ToInches());
            Assert.Equal(6, audit.PinX.ToInches(), 6);
            Assert.Equal(1, audit.PinY.ToInches(), 6);
            Assert.Equal(3, page.GetContainerMembers(runtime).Count);
            Assert.True(runtime.GetShapeBounds().Left.ToInches() < ingress.GetShapeBounds().Left.ToInches());
            Assert.True(runtime.GetShapeBounds().Right.ToInches() > worker.GetShapeBounds().Right.ToInches());
            Assert.NotEmpty(page.Connectors.Single(connector => connector.From?.Id == "ingress" && connector.To?.Id == "api").Waypoints);
            Assert.NotEmpty(page.Connectors.Single(connector => connector.From?.Id == "api" && connector.To?.Id == "worker").Waypoints);
            Assert.Empty(page.Connectors.Single(connector => connector.From?.Id == "worker" && connector.To?.Id == "audit").Waypoints);

            Assert.Empty(VisioValidator.Validate(updatedPath));
        }
    }
}
