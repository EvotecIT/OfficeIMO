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
    }
}
