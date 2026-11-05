using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioSwimlaneDiagramBuilderTests {
        [Fact]
        public void SwimlaneDiagramBuilderCreatesStyledProcessMapPage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SwimlaneDiagram("Order Fulfillment", swim => {
                    swim.Theme(VisioStyleTheme.Modern());
                    swim.Lane("customer", "Customer");
                    swim.Lane("sales", "Sales");
                    swim.Lane("ops", "Operations");
                    swim.Phase("request", "Request");
                    swim.Phase("review", "Review");
                    swim.Phase("approval", "Approval");
                    swim.Phase("fulfill", "Fulfill");
                    swim.Start("start", "Submit order", "customer", "request");
                    swim.Step("qualify", "Qualify order", "sales", "review");
                    swim.Decision("approved", "Approved?", "sales", "approval");
                    swim.Step("revise", "Revise request", "customer", "approval");
                    swim.Step("pick", "Pick items", "ops", "approval");
                    swim.Data("invoice", "Create invoice", "sales", "fulfill");
                    swim.End("ship", "Ship order", "ops", "fulfill");
                    swim.Flow("start", "qualify", "handoff");
                    swim.Flow("qualify", "approved");
                    swim.Exception("approved", "revise", "no");
                    swim.Handoff("approved", "pick", "yes");
                    swim.Flow("pick", "invoice");
                    swim.Flow("invoice", "ship");
                });

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Order Fulfillment", page.Name);
            Assert.Equal(17, page.Shapes.Count);
            Assert.Equal(6, page.Connectors.Count);
            Assert.Contains(page.Shapes, shape => shape.Id == "lane-customer" && shape.NameU == "Rectangle");
            Assert.Contains(page.Shapes, shape => shape.Id == "phase-approval" && shape.NameU == "Rectangle");
            Assert.Contains(page.Shapes, shape => shape.Id == "start" && shape.NameU == "Ellipse");
            Assert.Contains(page.Shapes, shape => shape.Id == "approved" && shape.NameU == "Decision");
            Assert.Contains(page.Shapes, shape => shape.Id == "invoice" && shape.NameU == "Data");
            VisioStencilProfile profile = document.CreateStencilProfile();
            Assert.Equal(14, profile.StencilBackedShapeCount);
            Assert.Equal(new[] { "Swimlane" }, profile.StencilCatalogs);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "swim.lane" && usage.Count == 3);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "swim.phase" && usage.Count == 4);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "swim.activity" && usage.Count == 3);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "swim.start-end" && usage.Count == 2);
            Assert.All(page.Connectors, connector => Assert.NotEmpty(connector.Waypoints));
            Assert.Contains(page.Connectors, connector => connector.Label == "yes" && connector.LinePattern == 1);
            Assert.Contains(page.Connectors, connector => connector.Label == "no" && connector.LinePattern == 2);
            Assert.Empty(page.AnalyzeVisualQuality().Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(17, loaded.Pages[0].Shapes.Count);
            Assert.Equal(6, loaded.Pages[0].Connectors.Count);
        }

        [Fact]
        public void SwimlaneDiagramBuilderPersistsTypedLanePhaseAndActivityMetadata() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SwimlaneDiagram("Fulfillment", swim => {
                    swim.Lane("sales", "Sales");
                    swim.Lane("ops", "Operations");
                    swim.Phase("review", "Review");
                    swim.Phase("fulfill", "Fulfill");
                    swim.Start("start", "Start", "sales", "review");
                    swim.Step("pick", "Pick", "ops", "fulfill");
                    swim.End("ship", "Ship", "ops", "fulfill");
                    swim.Flow("start", "pick");
                    swim.Flow("pick", "ship");
                });

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(2, page.GetSwimlaneLanes().Count);
            Assert.Equal(2, page.GetSwimlanePhases().Count);
            VisioSwimlaneActivityPlacement pick = Assert.Single(page.GetSwimlaneActivities(), activity => activity.Shape.Id == "pick");
            Assert.Equal("ops", pick.LaneId);
            Assert.Equal("fulfill", pick.PhaseId);
            Assert.Equal(VisioSwimlaneActivityKind.Step, pick.ActivityKind);
            Assert.Equal(VisioSemanticUserCells.SwimlaneLaneKind, page.Shapes.Single(shape => shape.Id == "lane-ops").GetUserCellValue(VisioSemanticUserCells.Kind));
            Assert.True(page.Shapes.Single(shape => shape.Id == "lane-ops").IsBackgroundSurface);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages[0];
            VisioSwimlaneActivityPlacement loadedPick = Assert.Single(loadedPage.GetSwimlaneActivities(), activity => activity.Shape.Id == "pick");
            Assert.Equal("ops", loadedPick.LaneId);
            Assert.Equal("fulfill", loadedPick.PhaseId);
            Assert.Equal(VisioSwimlaneActivityKind.Step, loadedPick.ActivityKind);
            Assert.Contains("OfficeIMO.SwimlaneLaneId", ReadZipText(filePath, "visio/pages/page1.xml"));
            Assert.Contains("OfficeIMO.SwimlanePhaseId", ReadZipText(filePath, "visio/pages/page1.xml"));
            Assert.Contains("OfficeIMO.SwimlaneActivityType", ReadZipText(filePath, "visio/pages/page1.xml"));
        }

        [Fact]
        public void LoadedSwimlaneActivitiesCanMoveWithFluentRelayoutAndSave() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string editedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            using var nativeDocument1 = VisioDocument.Create(filePath);
            nativeDocument1.SwimlaneDiagram("Fulfillment", swim => {
                swim.Lane("sales", "Sales");
                swim.Lane("ops", "Operations");
                swim.Phase("review", "Review");
                swim.Phase("fulfill", "Fulfill");
                swim.Step("qualify", "Qualify", "sales", "review");
                swim.Step("pick", "Pick", "ops", "fulfill");
                swim.End("ship", "Ship", "ops", "fulfill");
                swim.Flow("qualify", "pick");
                swim.Flow("pick", "ship");
            });
            nativeDocument1.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            loaded
                .GetPage("Fulfillment", page => { page.MoveSwimlaneActivity("pick", "sales", "review", options => {
                        options.AvoidShapes = false;
                        options.ActivityGap = 0.18D;
                    }); });

            VisioPage page = loaded.Pages[0];
            VisioShape pick = page.Shapes.Single(shape => shape.Id == "pick");
            VisioShape qualify = page.Shapes.Single(shape => shape.Id == "qualify");
            VisioSwimlaneLane? sales = page.FindSwimlaneLane("sales");
            VisioSwimlanePhase? review = page.FindSwimlanePhase("review");
            Assert.NotNull(sales);
            Assert.NotNull(review);
            OfficeIMO.Visio.VisioShapeBounds salesBounds = sales!.Body.GetShapeBounds();
            Assert.Equal("sales", pick.GetUserCellValue(VisioSemanticUserCells.SwimlaneLaneId));
            Assert.Equal("review", pick.GetUserCellValue(VisioSemanticUserCells.SwimlanePhaseId));
            Assert.Equal(review!.Header.PinX.ToInches(), pick.PinX.ToInches(), 6);
            Assert.InRange(pick.PinY.ToInches(), salesBounds.Bottom.ToInches(), salesBounds.Top.ToInches());
            Assert.InRange(qualify.PinY.ToInches(), salesBounds.Bottom.ToInches(), salesBounds.Top.ToInches());
            Assert.NotEqual(qualify.PinY.ToInches(), pick.PinY.ToInches());
            Assert.All(page.Connectors, connector => Assert.NotEmpty(connector.Waypoints));

            loaded.Save(editedPath);
            Assert.Empty(VisioValidator.Validate(editedPath));

            VisioDocument reloaded = VisioDocument.Load(editedPath);
            VisioSwimlaneActivityPlacement moved = Assert.Single(reloaded.Pages[0].GetSwimlaneActivities(), activity => activity.Shape.Id == "pick");
            Assert.Equal("sales", moved.LaneId);
            Assert.Equal("review", moved.PhaseId);
            Assert.Contains("OfficeIMO.SwimlaneLaneId", ReadZipText(editedPath, "visio/pages/page1.xml"));
        }

        [Fact]
        public void SwimlaneStencilCatalogExposesProcessMapShapes() {
            Assert.Equal("Swimlane", VisioStencils.Swimlane.Name);
            Assert.Equal("Activity", VisioStencils.Swimlane.Get("task").Name);
            Assert.Equal("Phase", VisioStencils.Swimlane.Get("milestone").Name);
            Assert.Equal("Start/End", VisioStencils.All.Get("swim.start-end").Name);
        }

        [Fact]
        public void SwimlaneDiagramBuilderStacksActivitiesInTheSameCell() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .SwimlaneDiagram("Packed Cell", swim => {
                    swim.Lane("ops", "Operations");
                    swim.Phase("work", "Work");
                    swim.Start("start", "Start", "ops", "work");
                    swim.Step("work", "Do work", "ops", "work");
                    swim.End("done", "Done", "ops", "work");
                });

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(6, page.Shapes.Count);
            Assert.True(page.FindShapeById("start")!.PinY.ToInches() > page.FindShapeById("work")!.PinY.ToInches());
            Assert.True(page.FindShapeById("work")!.PinY.ToInches() > page.FindShapeById("done")!.PinY.ToInches());
            Assert.Empty(page.AnalyzeVisualQuality().Select(issue => issue.ToString()));
        }

        [Fact]
        public void SwimlaneDiagramBuilderCanAddTitleWithoutOverlappingGrid() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SwimlaneDiagram("Order Fulfillment", swim => {
                    swim.Title();
                    swim.Lane("customer", "Customer");
                    swim.Lane("sales", "Sales");
                    swim.Phase("request", "Request");
                    swim.Phase("review", "Review");
                    swim.Start("start", "Submit order", "customer", "request");
                    swim.Step("qualify", "Qualify order", "sales", "review");
                    swim.Flow("start", "qualify", "handoff");
                });

            VisioPage page = Assert.Single(document.Pages);
            VisioShape title = Assert.Single(page.Shapes, shape => shape.Id == "title");
            VisioShape phaseHeader = Assert.Single(page.Shapes, shape => shape.Id == "phase-request");
            Assert.Equal("Text Box", title.NameU);
            Assert.Equal("Order Fulfillment", title.Text);
            Assert.True(title.PinY.ToInches() > phaseHeader.PinY.ToInches());
            Assert.Empty(page.AnalyzeVisualQuality().Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void SwimlaneDiagramBuilderCanAddSemanticCallouts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SwimlaneDiagram("Annotated Fulfillment", swim => {
                    swim.Title();
                    swim.Lane("customer", "Customer");
                    swim.Lane("sales", "Sales");
                    swim.Lane("ops", "Operations");
                    swim.Phase("request", "Request");
                    swim.Phase("approval", "Approval");
                    swim.Start("start", "Submit order", "customer", "request");
                    swim.Decision("approved", "Approved?", "sales", "approval");
                    swim.Step("pick", "Pick items", "ops", "approval");
                    swim.Flow("start", "approved");
                    swim.Handoff("approved", "pick", "yes");
                    swim.Callout("approved", "approval-note", "Escalate exceptions before fulfillment", 7.8, 5.9, options => {
                        options.Width = 2.8;
                        options.Height = 0.72;
                        options.RouteOffset = 0.12;
                    }); });

            VisioPage page = Assert.Single(document.Pages);
            VisioShape callout = Assert.Single(page.Callouts());
            VisioShape target = Assert.Single(page.Shapes, shape => shape.Id == "approved");
            Assert.Equal("approval-note", callout.Id);
            Assert.Equal("Escalate exceptions before fulfillment", callout.Text);
            Assert.Equal(target.Id, callout.CalloutTargetId);
            Assert.Contains("Annotations", callout.LayerNames);
            Assert.Equal(2.8, callout.Width.ToInches());
            Assert.Equal(0.72, callout.Height.ToInches());

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, callout));
            Assert.Same(target, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);
            Assert.Contains("Annotations", leader.LayerNames);
            Assert.Equal(leader.Id, callout.GetUserCellValue("OfficeIMO.CalloutLeaderId"));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void SwimlaneDiagramBuilderCanAutoPlaceSemanticCalloutsBesideActivities() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SwimlaneDiagram("Auto Annotated Fulfillment", swim => {
                    swim.Title();
                    swim.Lane("sales", "Sales");
                    swim.Lane("ops", "Operations");
                    swim.Phase("intake", "Intake");
                    swim.Phase("fulfill", "Fulfill");
                    swim.Start("qualify", "Qualify request", "sales", "intake");
                    swim.Decision("approved", "Approved?", "sales", "fulfill");
                    swim.End("ship", "Ship order", "ops", "fulfill");
                    swim.Flow("qualify", "approved");
                    swim.Handoff("approved", "ship");
                    swim.Callout("approved", "approval-note", "Escalate exceptions before fulfillment", VisioSide.Right, 0.45, options => {
                        options.Width = 2.65;
                        options.Height = 0.72;
                    }); swim.Callout("qualify", "Confirm account owner", VisioSide.Left, 0.25); });

            VisioPage page = Assert.Single(document.Pages);
            VisioShape approved = Assert.Single(page.Shapes, shape => shape.Id == "approved");
            VisioShape qualify = Assert.Single(page.Shapes, shape => shape.Id == "qualify");
            VisioShape explicitCallout = Assert.Single(page.Callouts(), shape => shape.Id == "approval-note");
            VisioShape generatedCallout = Assert.Single(page.Callouts(), shape => shape.Id == "qualify-callout");

            Assert.True(explicitCallout.PinX.ToInches() > approved.PinX.ToInches());
            Assert.Equal(approved.PinY.ToInches(), explicitCallout.PinY.ToInches(), 6);
            Assert.Equal(approved.Id, explicitCallout.CalloutTargetId);
            Assert.Equal(2.65, explicitCallout.Width.ToInches());
            Assert.True(generatedCallout.PinX.ToInches() < qualify.PinX.ToInches());
            Assert.Equal(qualify.Id, generatedCallout.CalloutTargetId);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, explicitCallout));
            Assert.Same(approved, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void SwimlaneDiagramBuilderKeepsBandsAndCalloutsInMetricPageUnits() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .SwimlaneDiagram("Metric Swimlane", swim => {
                    swim.PageSize(20, 12, VisioMeasurementUnit.Centimeters);
                    swim.GridSize(4, 2, 2, 1);
                    swim.Lane("sales", "Sales");
                    swim.Phase("review", "Review");
                    swim.Step("qualify", "Qualify", "sales", "review");
                    swim.Callout("qualify", "qualify-note", "Metric note", 9, 7, options => {
                        options.Width = 3;
                        options.Height = 1;
                    }); });

            VisioPage page = Assert.Single(document.Pages);
            VisioShape laneHeader = Assert.Single(page.Shapes, shape => shape.Id == "lane-header-sales");
            VisioShape activity = Assert.Single(page.Shapes, shape => shape.Id == "qualify");
            VisioShape callout = Assert.Single(page.Callouts());

            Assert.Equal(2D.ToInches(VisioMeasurementUnit.Centimeters), laneHeader.Width.ToInches(), 6);
            Assert.Equal(2D.ToInches(VisioMeasurementUnit.Centimeters), laneHeader.Height.ToInches(), 6);
            Assert.InRange(activity.PinX.ToInches(), 0D, page.Width.ToInches());
            Assert.Equal(9D.ToInches(VisioMeasurementUnit.Centimeters), callout.PinX.ToInches(), 6);
            Assert.Equal(7D.ToInches(VisioMeasurementUnit.Centimeters), callout.PinY.ToInches(), 6);
        }

        [Fact]
        public void SwimlaneDiagramBuilderGeneratesUniqueCalloutIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .SwimlaneDiagram("Generated", swim => {
                    swim.Lane("ops", "Operations");
                    swim.Phase("work", "Work");
                    swim.Step("review", "Review", "ops", "work");
                    swim.Callout("review", "First note", 5.3, 5.8);
                    swim.Callout("review", "Second note", 5.3, 4.9);
                });

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(new[] { "review-callout", "review-callout-2" }, page.Callouts().Select(shape => shape.Id).ToArray());
        }

        [Fact]
        public void SwimlaneDiagramBuilderRejectsUnknownFlowEndpoints() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.SwimlaneDiagram("Invalid", swim => { swim.Lane("sales", "Sales"); swim.Phase("review", "Review"); swim.Step("qualify", "Qualify", "sales", "review"); swim.Flow("qualify", "missing"); }));

            Assert.Contains("Unknown swimlane activity id", exception.Message);
        }

        [Fact]
        public void SwimlaneDiagramBuilderRejectsTitleIdCollisions() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.SwimlaneDiagram("Invalid", swim => { swim.Lane("sales", "Sales"); swim.Phase("review", "Review"); swim.Step("title", "Qualify", "sales", "review"); swim.Title(); }));

            Assert.Contains("already exists", exception.Message);
        }

        [Fact]
        public void SwimlaneDiagramBuilderRejectsCalloutIdCollisionsAndUnknownTargets() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknownTarget = Assert.Throws<ArgumentException>(() =>
                document.SwimlaneDiagram("Invalid", swim => {
                    swim.Lane("sales", "Sales");
                    swim.Phase("review", "Review");
                    swim.Step("qualify", "Qualify", "sales", "review");
                    swim.Callout("missing", "note", "No target", 4, 4);
                }));
            ArgumentException activityCollision = Assert.Throws<ArgumentException>(() =>
                document.SwimlaneDiagram("Invalid", swim => {
                    swim.Lane("sales", "Sales");
                    swim.Phase("review", "Review");
                    swim.Step("qualify", "Qualify", "sales", "review");
                    swim.Callout("qualify", "qualify", "Duplicate id", 4, 4);
                }));
            ArgumentException laneCollision = Assert.Throws<ArgumentException>(() =>
                document.SwimlaneDiagram("Invalid", swim => {
                    swim.Lane("sales", "Sales");
                    swim.Phase("review", "Review");
                    swim.Step("qualify", "Qualify", "sales", "review");
                    swim.Callout("qualify", "lane-sales", "Duplicate id", 4, 4);
                }));

            Assert.Contains("Unknown swimlane activity id", unknownTarget.Message);
            Assert.Contains("already exists", activityCollision.Message);
            Assert.Contains("already exists", laneCollision.Message);
        }

        [Fact]
        public void SwimlaneDiagramBuilderRejectsAutoCalloutPlacementIssues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentOutOfRangeException autoPlacement = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.SwimlaneDiagram("Invalid", swim => {
                    swim.Lane("ops", "Operations");
                    swim.Phase("review", "Review");
                    swim.Step("review", "Review", "ops", "review");
                    swim.Callout("review", "Invalid", VisioSide.Auto);
                }));
            ArgumentOutOfRangeException badGap = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.SwimlaneDiagram("Invalid", swim => {
                    swim.Lane("ops", "Operations");
                    swim.Phase("review", "Review");
                    swim.Step("review", "Review", "ops", "review");
                    swim.Callout("review", "Invalid", VisioSide.Right, double.NaN);
                }));

            Assert.Contains("Placement must be", autoPlacement.Message);
            Assert.Contains("finite non-negative", badGap.Message);
        }

        private static string ReadZipText(string filePath, string entryName) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            using Stream stream = archive.GetEntry(entryName)!.Open();
            using StreamReader reader = new(stream);
            return reader.ReadToEnd();
        }
    }
}
