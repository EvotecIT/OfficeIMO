using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioTimelineDiagramBuilderTests {
        [Fact]
        public void TimelineDiagramBuilderTracksLargeUniqueIdSetsWithoutRescanningItems() {
            var builder = new VisioTimelineDiagramBuilder(VisioDocument.Create(), "Security Timeline");
            DateTime date = new DateTime(2026, 1, 1);

            for (int index = 0; index < 4096; index++) {
                builder.Milestone("milestone-" + index, date, "Milestone " + index);
            }

            Assert.Throws<ArgumentException>(() => builder.Span(
                "milestone-2048",
                date,
                date.AddDays(1),
                "Duplicate"));
        }

        [Fact]
        public void TimelineDiagramBuilderReleasesReplacedTitleId() {
            var builder = new VisioTimelineDiagramBuilder(VisioDocument.Create(), "Security Timeline");
            DateTime date = new DateTime(2026, 1, 1);

            builder
                .Title("Initial", id: "old-title")
                .Title("Replacement", id: "new-title")
                .Milestone("old-title", date, "Reused id");
        }

        [Fact]
        public void TimelineDiagramBuilderCreatesStyledRoadmapPage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .TimelineDiagram("Product Roadmap", timeline => timeline
                    .Theme(VisioStyleTheme.Modern())
                    .Range(new DateTime(2026, 1, 1), new DateTime(2026, 6, 30))
                    .Span("discovery", new DateTime(2026, 1, 8), new DateTime(2026, 2, 20), "Discovery", 0)
                    .Span("build", new DateTime(2026, 2, 21), new DateTime(2026, 5, 15), "Build", 1)
                    .Span("enablement", new DateTime(2026, 4, 1), new DateTime(2026, 6, 10), "Enablement", 0, VisioTimelinePlacement.Below)
                    .Milestone("kickoff", new DateTime(2026, 1, 12), "Kickoff", VisioTimelinePlacement.Above)
                    .Decision("gate", new DateTime(2026, 2, 25), "Go / no-go", VisioTimelinePlacement.Below)
                    .Risk("risk", new DateTime(2026, 3, 18), "Security review", VisioTimelinePlacement.Above)
                    .Release("preview", new DateTime(2026, 5, 20), "Public preview", VisioTimelinePlacement.Below)
                    .Milestone("ga", new DateTime(2026, 6, 25), "GA", VisioTimelinePlacement.Above));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Product Roadmap", page.Name);
            Assert.Equal(16, page.Shapes.Count);
            Assert.Empty(page.Connectors);
            Assert.Contains(page.Shapes, shape => shape.Id == "timeline-axis" && shape.NameU == "Rectangle");
            Assert.Contains(page.Shapes, shape => shape.Id == "kickoff" && shape.NameU == "Diamond");
            Assert.Contains(page.Shapes, shape => shape.Id == "preview" && shape.NameU == "Circle");
            VisioStencilProfile profile = document.CreateStencilProfile();
            Assert.Equal(16, profile.StencilBackedShapeCount);
            Assert.Equal(new[] { "Timeline" }, profile.StencilCatalogs);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "time.axis" && usage.Count == 1);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "time.span" && usage.Count == 3);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "time.label" && usage.Count == 7);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "time.milestone" && usage.Count == 2);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "time.decision" && usage.Count == 1);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "time.risk" && usage.Count == 1);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "time.release" && usage.Count == 1);
            Assert.True(page.FindShapeById("kickoff")!.PinX < page.FindShapeById("ga")!.PinX);
            Assert.True(page.FindShapeById("discovery")!.Width < page.FindShapeById("build")!.Width);
            string[] qualityIssues = page.AnalyzeVisualQuality().Select(issue => issue.ToString()).ToArray();
            Assert.True(qualityIssues.Length == 0, string.Join(Environment.NewLine, qualityIssues));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(16, loaded.Pages[0].Shapes.Count);
            Assert.Empty(loaded.Pages[0].Connectors);
        }

        [Fact]
        public void TimelineStencilCatalogExposesRoadmapShapes() {
            Assert.Equal("Timeline", VisioStencils.Timeline.Name);
            Assert.Equal("Milestone", VisioStencils.Timeline.Get("marker").Name);
            Assert.Equal("Span", VisioStencils.Timeline.Get("duration").Name);
            Assert.Equal("Risk", VisioStencils.All.Get("time.risk").Name);
        }

        [Fact]
        public void TimelineDiagramBuilderCanAddTitleWithoutOverlappingRoadmap() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .TimelineDiagram("Product Roadmap", timeline => timeline
                    .Title()
                    .PageSize(6, 3)
                    .Range(new DateTime(2026, 1, 1), new DateTime(2026, 3, 31))
                    .Span("build", new DateTime(2026, 1, 15), new DateTime(2026, 3, 10), "Build")
                    .Milestone("kickoff", new DateTime(2026, 1, 20), "Kickoff", VisioTimelinePlacement.Above)
                    .Release("preview", new DateTime(2026, 3, 15), "Preview", VisioTimelinePlacement.Below));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape title = Assert.Single(page.Shapes, shape => shape.Id == "title");
            double highestRoadmapTop = page.Shapes
                .Where(shape => shape.Id != "title")
                .Max(shape => shape.PinY + shape.Height / 2D);
            Assert.Equal("Text Box", title.NameU);
            Assert.Equal("Product Roadmap", title.Text);
            Assert.True(title.PinY - title.Height / 2D > highestRoadmapTop);
            Assert.Contains(page.Shapes, shape => shape.Id == "kickoff-label");
            Assert.Empty(page.AnalyzeVisualQuality().Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void TimelineDiagramBuilderCanAddSemanticCalloutsToMilestonesAndSpans() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .TimelineDiagram("Annotated Roadmap", timeline => timeline
                    .Title()
                    .Range(new DateTime(2026, 1, 1), new DateTime(2026, 6, 30))
                    .Span("build", new DateTime(2026, 2, 1), new DateTime(2026, 5, 15), "Build", 0)
                    .Risk("risk", new DateTime(2026, 3, 18), "Security review", VisioTimelinePlacement.Above)
                    .Release("preview", new DateTime(2026, 5, 20), "Public preview", VisioTimelinePlacement.Below)
                    .Callout(" risk ", "risk-note", "Resolve before preview", 6.7, 6.4, options => {
                        options.Width = 2.45;
                        options.Height = 0.72;
                        options.RouteOffset = 0.11;
                    })
                    .Callout("build", "build-note", "Implementation runway", 5.2, 5.7));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape[] callouts = page.Callouts().ToArray();
            Assert.Equal(2, callouts.Length);

            VisioShape riskCallout = Assert.Single(callouts, shape => shape.Id == "risk-note");
            VisioShape buildCallout = Assert.Single(callouts, shape => shape.Id == "build-note");
            VisioShape risk = Assert.Single(page.Shapes, shape => shape.Id == "risk");
            VisioShape build = Assert.Single(page.Shapes, shape => shape.Id == "build");

            Assert.Equal("Resolve before preview", riskCallout.Text);
            Assert.Equal(risk.Id, riskCallout.CalloutTargetId);
            Assert.Equal(build.Id, buildCallout.CalloutTargetId);
            Assert.Contains("Annotations", riskCallout.LayerNames);
            Assert.Equal(2.45, riskCallout.Width);
            Assert.Equal(0.72, riskCallout.Height);

            VisioConnector riskLeader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, riskCallout));
            VisioConnector buildLeader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, buildCallout));
            Assert.Same(risk, riskLeader.To);
            Assert.Same(build, buildLeader.To);
            Assert.Equal(EndArrow.None, riskLeader.EndArrow);
            Assert.Contains("Annotations", buildLeader.LayerNames);
            Assert.Equal(riskLeader.Id, riskCallout.GetUserCellValue("OfficeIMO.CalloutLeaderId"));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void TimelineDiagramBuilderCanAutoPlaceSemanticCalloutsBesideItems() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .TimelineDiagram("Auto Annotated Roadmap", timeline => timeline
                    .Title()
                    .Range(new DateTime(2026, 1, 1), new DateTime(2026, 6, 30))
                    .Span("build", new DateTime(2026, 2, 1), new DateTime(2026, 5, 15), "Build", 0)
                    .Risk("risk", new DateTime(2026, 3, 18), "Security review", VisioTimelinePlacement.Above)
                    .Release("preview", new DateTime(2026, 5, 20), "Public preview", VisioTimelinePlacement.Below)
                    .Callout("risk", "risk-note", "Resolve before preview", VisioSide.Top, 0.35, options => {
                        options.Width = 2.45;
                        options.Height = 0.72;
                    })
                    .Callout("build", "Implementation runway", VisioSide.Bottom, 0.25));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape risk = Assert.Single(page.Shapes, shape => shape.Id == "risk");
            VisioShape build = Assert.Single(page.Shapes, shape => shape.Id == "build");
            VisioShape explicitCallout = Assert.Single(page.Callouts(), shape => shape.Id == "risk-note");
            VisioShape generatedCallout = Assert.Single(page.Callouts(), shape => shape.Id == "build-callout");

            Assert.True(explicitCallout.PinY > risk.PinY);
            Assert.Equal(risk.PinX, explicitCallout.PinX, 6);
            Assert.Equal(risk.Id, explicitCallout.CalloutTargetId);
            Assert.Equal(2.45, explicitCallout.Width);
            Assert.True(generatedCallout.PinY < build.PinY);
            Assert.Equal(build.Id, generatedCallout.CalloutTargetId);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, explicitCallout));
            Assert.Same(risk, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void TimelineDiagramBuilderKeepsCalloutPinsInMetricPageUnits() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .TimelineDiagram("Metric Timeline", timeline => timeline
                    .PageSize(20, 12, VisioMeasurementUnit.Centimeters)
                    .Range(new DateTime(2026, 1, 1), new DateTime(2026, 3, 31))
                    .Milestone("kickoff", new DateTime(2026, 1, 15), "Kickoff")
                    .Callout("kickoff", "kickoff-note", "Metric note", 8, 7, options => {
                        options.Width = 3;
                        options.Height = 1;
                    }));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape callout = Assert.Single(page.Callouts());

            Assert.Equal(8D.ToInches(VisioMeasurementUnit.Centimeters), callout.PinX, 6);
            Assert.Equal(7D.ToInches(VisioMeasurementUnit.Centimeters), callout.PinY, 6);
            Assert.Equal(3D.ToInches(VisioMeasurementUnit.Centimeters), callout.Width, 6);
        }

        [Fact]
        public void TimelineDiagramBuilderGeneratesUniqueCalloutIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .TimelineDiagram("Generated", timeline => timeline
                    .Span("build", new DateTime(2026, 2, 1), new DateTime(2026, 3, 15), "Build")
                    .Callout("build", "First note", 5.3, 5.8)
                    .Callout("build", "Second note", 5.3, 4.9));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(new[] { "build-callout", "build-callout-2" }, page.Callouts().Select(shape => shape.Id).ToArray());
        }

        [Fact]
        public void TimelineDiagramBuilderRejectsItemsOutsideConfiguredRange() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.TimelineDiagram("Invalid", timeline => timeline
                    .Range(new DateTime(2026, 1, 1), new DateTime(2026, 1, 31))
                    .Milestone("late", new DateTime(2026, 2, 1), "Late")));
        }

        [Fact]
        public void TimelineDiagramBuilderRejectsGeneratedShapeIdCollisions() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException titleFirst = Assert.Throws<ArgumentException>(() =>
                document.TimelineDiagram("Invalid", timeline => timeline
                    .Title()
                    .Milestone("title", new DateTime(2026, 1, 1), "Kickoff")));
            ArgumentException itemFirst = Assert.Throws<ArgumentException>(() =>
                document.TimelineDiagram("Invalid", timeline => timeline
                    .Milestone("kickoff", new DateTime(2026, 1, 1), "Kickoff")
                    .Title(id: "kickoff-label")));
            ArgumentException axisCollision = Assert.Throws<ArgumentException>(() =>
                document.TimelineDiagram("Invalid", timeline => timeline
                    .Milestone("timeline-axis", new DateTime(2026, 1, 1), "Kickoff")));
            ArgumentException generatedLabelCollision = Assert.Throws<ArgumentException>(() =>
                document.TimelineDiagram("Invalid", timeline => timeline
                    .Milestone("kickoff", new DateTime(2026, 1, 1), "Kickoff")
                    .Span("kickoff-label", new DateTime(2026, 1, 2), new DateTime(2026, 1, 3), "Blocked")));

            Assert.Contains("already exists", titleFirst.Message);
            Assert.Contains("already exists", itemFirst.Message);
            Assert.Contains("already exists", axisCollision.Message);
            Assert.Contains("already exists", generatedLabelCollision.Message);
        }

        [Fact]
        public void TimelineDiagramBuilderRejectsCalloutIdCollisionsAndUnknownTargets() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknownTarget = Assert.Throws<ArgumentException>(() =>
                document.TimelineDiagram("Invalid", timeline => timeline
                    .Milestone("kickoff", new DateTime(2026, 1, 1), "Kickoff")
                    .Callout("missing", "note", "No target", 4, 4)));
            ArgumentException milestoneCollision = Assert.Throws<ArgumentException>(() =>
                document.TimelineDiagram("Invalid", timeline => timeline
                    .Milestone("kickoff", new DateTime(2026, 1, 1), "Kickoff")
                    .Callout("kickoff", "kickoff", "Duplicate id", 4, 4)));
            ArgumentException generatedLabelCollision = Assert.Throws<ArgumentException>(() =>
                document.TimelineDiagram("Invalid", timeline => timeline
                    .Milestone("kickoff", new DateTime(2026, 1, 1), "Kickoff")
                    .Callout("kickoff", "kickoff-label", "Duplicate id", 4, 4)));

            Assert.Contains("Unknown timeline item id", unknownTarget.Message);
            Assert.Contains("already exists", milestoneCollision.Message);
            Assert.Contains("already exists", generatedLabelCollision.Message);
        }

        [Fact]
        public void TimelineDiagramBuilderRejectsAutoCalloutPlacementIssues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentOutOfRangeException autoPlacement = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.TimelineDiagram("Invalid", timeline => timeline
                    .Milestone("kickoff", new DateTime(2026, 1, 1), "Kickoff")
                    .Callout("kickoff", "Invalid", VisioSide.Auto)));
            ArgumentOutOfRangeException badGap = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.TimelineDiagram("Invalid", timeline => timeline
                    .Milestone("kickoff", new DateTime(2026, 1, 1), "Kickoff")
                    .Callout("kickoff", "Invalid", VisioSide.Right, double.NaN)));

            Assert.Contains("Placement must be", autoPlacement.Message);
            Assert.Contains("finite non-negative", badGap.Message);
        }
    }
}
