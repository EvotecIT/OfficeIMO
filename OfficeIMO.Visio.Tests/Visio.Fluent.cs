using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.OfficeColor;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioDocumentTests {
        [Fact]
        public void NativePageCreationReturnsTheCreatedPage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            VisioPage result = document
                .AddPage("Page1", OfficeLength.Inches(8.5), OfficeLength.Inches(11), VisioMeasurementUnit.Inches, configure: p => { });

            Assert.Same(document.Pages[0], result);
            Assert.Single(document.Pages);
            Assert.Equal("Page1", document.Pages[0].Name);
            document.Save();
        }

        [Fact]
        public void NativeShapesUsePageUnitsAndPersistNumericShapeIds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document
                .AddPage("Metric", OfficeLength.Centimeters(10), OfficeLength.Centimeters(10), VisioMeasurementUnit.Centimeters, configure: p => p.AddRectangle("box", OfficeLength.Centimeters(2.54), OfficeLength.Centimeters(2.54), OfficeLength.Centimeters(2.54), OfficeLength.Centimeters(2.54), "Box"));
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
        public void NativeConnectCanTargetSidesAndStyleLines() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document
                .AddPage("Page1", p => {
                    p.AddRectangle("left", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Left");
                    p.AddRectangle("right", OfficeLength.Inches(5), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Right");
                    p.AddConnector("left", "right", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: c => {
                        c.SetKind(ConnectorKind.RightAngle);
                        c.SetLineColor(Color.DarkBlue);
                        c.SetArrowEnd(EndArrow.Triangle);
                    });
                });
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
        public void NativeConnectUsesPageScopedConnectorIdsAcrossDocuments() {
            static VisioDocument CreateDocument(string filePath) {
                VisioDocument document = VisioDocument.Create(filePath);
                document
                    .AddPage("Page1", p => {
                        p.AddRectangle("1", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Left");
                        p.AddRectangle("2", OfficeLength.Inches(5), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Right");
                        p.AddConnector("1", "2");
                    });
                return document;
            }

            VisioDocument first = CreateDocument(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioDocument second = CreateDocument(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            Assert.Equal("3", Assert.Single(first.Pages[0].Connectors).Id);
            Assert.Equal("3", Assert.Single(second.Pages[0].Connectors).Id);
        }

        [Fact]
        public void NativePageCanResolveShapeOverlaps() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document
                .AddPage("Page1", p => {
                    p.AddRectangle("first", OfficeLength.Inches(2), OfficeLength.Inches(3), OfficeLength.Inches(1.2), OfficeLength.Inches(0.8), "First");
                    p.AddRectangle("second", OfficeLength.Inches(2.2), OfficeLength.Inches(3), OfficeLength.Inches(1.2), OfficeLength.Inches(0.8), "Second");
                    p.AddRectangle("third", OfficeLength.Inches(2.4), OfficeLength.Inches(3), OfficeLength.Inches(1.2), OfficeLength.Inches(0.8), "Third");
                    p.ResolveShapeOverlaps(step: 0.25, maxAttempts: 12);
                });

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
        public void NativePageCanAddTitleAndTextBoxAdornments() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document
                .AddPage("Overview", p => {
                    p.AddTitle("System Overview");
                    p.AddTextBox("caption", OfficeLength.Inches(4.25), OfficeLength.Inches(9.8), OfficeLength.Inches(5.5), OfficeLength.Inches(0.35), "Generated by OfficeIMO", shape => {
                        shape.SetTextColor(Color.DarkBlue);
                        shape.SetFontSize(11);
                    });
                });

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

            document
                .AddPage("Page1", p => {
                    p.AddRectangle("left", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(2), "Left");
                    p.AddRectangle("right", OfficeLength.Inches(5), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(2), "Right");
                });

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
        public void NativeCanEditLoadedExistingPageWithoutAddingDuplicatePages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.AddPage("Operations", p => {
                p.AddRectangle("api", OfficeLength.Inches(2), OfficeLength.Inches(4), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "Legacy API");
                p.AddRectangle("db", OfficeLength.Inches(5), OfficeLength.Inches(4), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "Database");
                p.GetShape("api", shape => shape.SetShapeData("Owner", "Ops"));
                p.GetShape("db", shape => shape.SetShapeData("Owner", "Data"));
                p.AddConnector("api", "db", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => connector.SetLabel("SQL"));
            });
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            loaded.GetPage("Operations", page => {
                page.ShapesWithData("Owner", "Ops", selection => {
                    selection.Fill(Color.LightBlue);
                    selection.ShapeData("Reviewed", "Yes", "Reviewed", VisioShapeDataType.Boolean);
                });
                page.ShapesContainingText("Legacy", selection => selection.Text(shape => shape.Text!.Replace("Legacy", "Production")));
                page.AddRectangle("monitor", OfficeLength.Inches(7), OfficeLength.Inches(4), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "Monitor");
                page.AddConnector("api", "monitor", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => {
                    connector.SetKind(ConnectorKind.RightAngle);
                    connector.SetArrowEnd(EndArrow.Triangle);
                    connector.SetLabel("metrics");
                });
                page.ConfigureConnectors(selection => {
                    selection.LineColor(Color.DarkBlue);
                    selection.LabelPosition();
                });
            });
            loaded.Save(updatedPath);

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
        public void NativePageOrAddIsIdempotentForCreateOrEditWorkflows() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            document.GetOrAddPage("Inventory", page => page.AddRectangle("server", OfficeLength.Inches(2), OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(1), "Server"));
            document.GetOrAddPage("Inventory", page => {
                page.GetShape("server", shape => shape.SetShapeData("Owner", "Ops"));
            });

            VisioPage page = Assert.Single(document.Pages);
            VisioShape server = Assert.Single(page.Shapes);
            Assert.Equal("Inventory", page.Name);
            Assert.Equal("Ops", server.GetShapeDataValue("Owner"));
        }

        [Fact]
        public void NativeAdvancedSelectionsEditLoadedGeometryPathsAndConnectors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            using var nativeDocument6 = VisioDocument.Create(filePath);
            nativeDocument6.AddPage("Network", page => {
                page.AddRectangle("zone", OfficeLength.Inches(5), OfficeLength.Inches(4), OfficeLength.Inches(6), OfficeLength.Inches(4), "Runtime Zone");
                page.AddRectangle("ingress", OfficeLength.Inches(2.5), OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(1), "Ingress");
                page.AddRectangle("api", OfficeLength.Inches(4.5), OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(1), "API");
                page.AddRectangle("db", OfficeLength.Inches(6.5), OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(1), "Database");
                page.AddRectangle("outside", OfficeLength.Inches(9.5), OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(1), "Outside");
                page.GetShape("api", shape => {
                    shape.SetUserCell("Tier", "App");
                    shape.AddHyperlink("https://example.org/api");
                    shape.Protect(protection => protection.Position());
                });
                page.GetShape("db", shape => shape.SetLayer("Data"));
                page.AddConnector("ingress", "api", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => {
                    connector.SetKind(ConnectorKind.RightAngle);
                    connector.SetLayer("traffic");
                    connector.AddHyperlink("https://example.org/flow");
                    connector.SetLabel("https");
                });
                page.AddConnector("api", "db", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => {
                    connector.SetKind(ConnectorKind.RightAngle);
                    connector.SetLayer("traffic");
                    connector.Protect(protection => protection.Endpoints());
                    connector.SetLabel("sql");
                });
            });
            nativeDocument6.Save();

            using var nativeDocument5 = VisioDocument.Load(filePath);
            nativeDocument5.GetPage("Network", page => {
                page.ShapesContainedIn("zone", selection => {
                    selection.ShapeData("Zone", "Core");
                    selection.Fill(Color.LightGreen);
                });
                page.PathBetween("ingress", "db", selection => {
                    selection.ShapeData("CriticalPath", "Yes");
                    selection.Stroke(Color.Red, 0.03);
                });
                page.ConnectedComponent("ingress", selection => selection.ShapeData("Component", "Runtime"));
                page.ShapesWithUserCell("Tier", "App", selection => selection.ShapeData("TierReviewed", "Yes"));
                page.ShapesWithHyperlink("https://example.org/api", selection => selection.Fill(Color.LightBlue));
                page.ShapesWithProtection(protection => protection.LockMoveX == true, selection => selection.ShapeData("Locked", "Position"));
                page.ConnectedConnectors("api", selection => {
                    selection.LineColor(Color.DarkBlue);
                    selection.LabelPosition();
                });
                page.OutgoingConnectors("api", selection => selection.EndArrow(EndArrow.Triangle));
                page.ConnectorsInLayer("traffic", selection => selection.LineWeight(0.04));
                page.ConnectorsWithHyperlink("https://example.org/flow", selection => selection.Label("observed"));
                page.ConnectorsWithProtection(protection => protection.LockBegin == true, selection => selection.LinePattern(2));
            });
            nativeDocument5.Save(updatedPath);

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
        public void NativeReplaceMasterStandardizesLoadedStencilShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            using var nativeDocument4 = VisioDocument.Create(filePath);
            nativeDocument4.AddPage("Standardize", page => {
                page.AddStencilShape(id: "source", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(2), y: OfficeLength.Inches(5), text: "Source");
                page.AddStencilShape(id: "task", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(5), y: OfficeLength.Inches(5), text: "Review");
                page.AddStencilShape(id: "archive", stencil: VisioStencils.Flowchart.Get("data"), x: OfficeLength.Inches(8), y: OfficeLength.Inches(5), text: "Archive");
                page.GetShape("task", shape => {
                    shape.SetFill(Color.LightYellow);
                    shape.SetStroke(Color.Orange, 0.02);
                    shape.SetLayer("Review");
                    shape.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String, "Owning team");
                    shape.SetUserCell("Stage", "Review", "STR");
                    shape.AddHyperlink("https://example.org/review", "Review docs");
                });
                page.AddConnector("source", "task", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => connector.SetLabel("submit"));
                page.AddConnector("task", "archive", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => connector.SetLabel("store"));
            });
            nativeDocument4.Save();

            using var nativeDocument3 = VisioDocument.Load(filePath);
            nativeDocument3.GetPage("Standardize", page => {
                page.ReplaceMaster("task", VisioStencils.Flowchart.Get("decision"), resizeToMaster: true);
                page.ReplaceMastersByMaster("Process", VisioStencils.Flowchart.Get("preparation"), resizeToMaster: true);
            });
            nativeDocument3.Save(updatedPath);

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
        public void NativeCanRelayoutLoadedContainerMembersAndPersistRoutedConnectors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            using var nativeDocument2 = VisioDocument.Create(filePath);
            nativeDocument2.AddPage("Operations", page => {
                page.AddRectangle("ingress", OfficeLength.Inches(7), OfficeLength.Inches(2), OfficeLength.Inches(1.2), OfficeLength.Inches(0.7), "Ingress");
                page.AddRectangle("api", OfficeLength.Inches(3), OfficeLength.Inches(6), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "API");
                page.AddRectangle("worker", OfficeLength.Inches(8), OfficeLength.Inches(5), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "Worker");
                page.AddRectangle("audit", OfficeLength.Inches(6), OfficeLength.Inches(1), OfficeLength.Inches(1.2), OfficeLength.Inches(0.7), "Audit");
                page.AddContainer("runtime", "Runtime", new[] {
                    "ingress", "api", "worker"
                }, options => {
                        options.Margin = 0.2D;
                        options.HeadingHeight = 0.3D;
                    });
                    page.AddConnector("ingress", "api", fromSide: VisioSide.Top, toSide: VisioSide.Bottom, configure: connector => connector.SetLabel("https"));
                    page.AddConnector("api", "worker", fromSide: VisioSide.Right, toSide: VisioSide.Left, configure: connector => connector.SetLabel("queue"));
                    page.AddConnector("worker", "audit", fromSide: VisioSide.Bottom, toSide: VisioSide.Top, configure: connector => connector.SetLabel("events"));
                    });
            nativeDocument2.Save();

            using var nativeDocument1 = VisioDocument.Load(filePath);
            nativeDocument1.GetPage("Operations", page => { page.RelayoutContainerMembers("runtime", options => {
                        options.Columns = 1;
                        options.HorizontalSpacing = 0D;
                        options.VerticalSpacing = 0.25D;
                        options.Order = VisioSelectionLayoutOrder.TopLeftToBottomRight;
                    }); page.AlignShapes(VisioHorizontalAlignment.Center, new[] { "ingress", "api", "worker" }); });
            nativeDocument1.Save(updatedPath);

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
