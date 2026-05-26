using System;
using System.IO;
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
                .SwimlaneDiagram("Order Fulfillment", swim => swim
                    .Theme(VisioStyleTheme.Modern())
                    .Lane("customer", "Customer")
                    .Lane("sales", "Sales")
                    .Lane("ops", "Operations")
                    .Phase("request", "Request")
                    .Phase("review", "Review")
                    .Phase("approval", "Approval")
                    .Phase("fulfill", "Fulfill")
                    .Start("start", "Submit order", "customer", "request")
                    .Step("qualify", "Qualify order", "sales", "review")
                    .Decision("approved", "Approved?", "sales", "approval")
                    .Step("revise", "Revise request", "customer", "approval")
                    .Step("pick", "Pick items", "ops", "approval")
                    .Data("invoice", "Create invoice", "sales", "fulfill")
                    .End("ship", "Ship order", "ops", "fulfill")
                    .Flow("start", "qualify", "handoff")
                    .Flow("qualify", "approved")
                    .Exception("approved", "revise", "no")
                    .Handoff("approved", "pick", "yes")
                    .Flow("pick", "invoice")
                    .Flow("invoice", "ship"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Order Fulfillment", page.Name);
            Assert.Equal(17, page.Shapes.Count);
            Assert.Equal(6, page.Connectors.Count);
            Assert.Contains(page.Shapes, shape => shape.Id == "lane-customer" && shape.NameU == "Rectangle");
            Assert.Contains(page.Shapes, shape => shape.Id == "phase-approval" && shape.NameU == "Rectangle");
            Assert.Contains(page.Shapes, shape => shape.Id == "start" && shape.NameU == "Ellipse");
            Assert.Contains(page.Shapes, shape => shape.Id == "approved" && shape.NameU == "Decision");
            Assert.Contains(page.Shapes, shape => shape.Id == "invoice" && shape.NameU == "Data");
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
        public void SwimlaneStencilCatalogExposesProcessMapShapes() {
            Assert.Equal("Swimlane", VisioStencils.Swimlane.Name);
            Assert.Equal("Activity", VisioStencils.Swimlane.Get("task").Name);
            Assert.Equal("Phase", VisioStencils.Swimlane.Get("milestone").Name);
            Assert.Equal("Start/End", VisioStencils.All.Get("swim.start-end").Name);
        }

        [Fact]
        public void SwimlaneDiagramBuilderStacksActivitiesInTheSameCell() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .SwimlaneDiagram("Packed Cell", swim => swim
                    .Lane("ops", "Operations")
                    .Phase("work", "Work")
                    .Start("start", "Start", "ops", "work")
                    .Step("work", "Do work", "ops", "work")
                    .End("done", "Done", "ops", "work"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(6, page.Shapes.Count);
            Assert.True(page.FindShapeById("start")!.PinY > page.FindShapeById("work")!.PinY);
            Assert.True(page.FindShapeById("work")!.PinY > page.FindShapeById("done")!.PinY);
            Assert.Empty(page.AnalyzeVisualQuality().Select(issue => issue.ToString()));
        }

        [Fact]
        public void SwimlaneDiagramBuilderRejectsUnknownFlowEndpoints() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.SwimlaneDiagram("Invalid", swim => swim
                    .Lane("sales", "Sales")
                    .Phase("review", "Review")
                    .Step("qualify", "Qualify", "sales", "review")
                    .Flow("qualify", "missing")));

            Assert.Contains("Unknown swimlane activity id", exception.Message);
        }
    }
}
