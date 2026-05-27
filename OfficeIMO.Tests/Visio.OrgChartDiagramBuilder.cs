using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioOrgChartDiagramBuilderTests {
        [Fact]
        public void OrgChartDiagramBuilderCreatesStyledHierarchyPage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .OrgChartDiagram("Leadership", org => org
                    .Theme(VisioStyleTheme.Modern())
                    .Root("ceo", "Marta Nowak", "Chief Executive Officer")
                    .Assistant("ea", "Eli Green", "Executive Assistant", "ceo")
                    .Manager("cto", "Alex Chen", "Chief Technology Officer", "ceo")
                    .Manager("coo", "Sam Rivera", "Chief Operating Officer", "ceo")
                    .Manager("cfo", "Priya Shah", "Chief Financial Officer", "ceo")
                    .TeamBand("engineering", "Engineering", "cto")
                    .TeamBand("operations", "Operations", "coo")
                    .Position("platform", "Nina Patel", "Platform Lead", "cto", "engineering")
                    .Position("security", "Owen Brooks", "Security Lead", "cto", "engineering")
                    .Vacancy("sre", "Open SRE Role", "coo", "operations")
                    .External("advisor", "Taylor Reed", "Advisor", "cfo"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Leadership", page.Name);
            Assert.Equal(11, page.Shapes.Count);
            Assert.Equal(8, page.Connectors.Count);
            Assert.Contains(page.Shapes, shape => shape.Id == "ceo" && shape.NameU == "Process");
            Assert.Contains(page.Shapes, shape => shape.Id == "ea" && shape.NameU == "Rectangle");
            Assert.Contains(page.Shapes, shape => shape.Id == "org-band-engineering" && shape.NameU == "Rectangle");
            Assert.Contains(page.Shapes, shape => shape.Id == "sre" && shape.NameU == "Rectangle");
            Assert.All(page.Connectors, connector => Assert.NotEmpty(connector.Waypoints));
            Assert.Empty(page.AnalyzeVisualQuality().Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(11, loaded.Pages[0].Shapes.Count);
            Assert.Equal(8, loaded.Pages[0].Connectors.Count);
        }

        [Fact]
        public void OrgChartStencilCatalogExposesCommonOrgChartShapes() {
            Assert.Equal("Org Chart", VisioStencils.OrgChart.Name);
            Assert.Equal("Executive", VisioStencils.OrgChart.Get("ceo").Name);
            Assert.Equal("Assistant", VisioStencils.OrgChart.Get("staff").Name);
            Assert.Equal("Team Band", VisioStencils.All.Get("org.team-band").Name);
        }

        [Fact]
        public void OrgChartDiagramBuilderCanAddTitleWithoutOverlappingHierarchy() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .OrgChartDiagram("Leadership", org => org
                    .Title()
                    .Root("ceo", "Marta Nowak", "Chief Executive Officer")
                    .Assistant("ea", "Eli Green", "Executive Assistant", " ceo ")
                    .TeamBand("engineering", "Engineering", " ceo ")
                    .Position("platform", "Nina Patel", "Platform Lead", " ceo ", " engineering "));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape title = Assert.Single(page.Shapes, shape => shape.Id == "title");
            double highestChartTop = page.Shapes
                .Where(shape => shape.Id != "title")
                .Max(shape => shape.PinY + shape.Height / 2D);
            Assert.Equal("Text Box", title.NameU);
            Assert.Equal("Leadership", title.Text);
            Assert.True(title.PinY - title.Height / 2D > highestChartTop);
            Assert.Contains(page.Connectors, connector => connector.From.Id == "ceo" && connector.To.Id == "ea");
            Assert.Contains(page.Shapes, shape => shape.Id == "org-band-engineering");
            Assert.Empty(page.AnalyzeVisualQuality().Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void OrgChartDiagramBuilderCanAddSemanticCallouts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .OrgChartDiagram("Annotated Leadership", org => org
                    .Title()
                    .Root("ceo", "Marta Nowak", "Chief Executive Officer")
                    .Assistant("ea", "Eli Green", "Executive Assistant", "ceo")
                    .Manager("cto", "Alex Chen", "Chief Technology Officer", "ceo")
                    .TeamBand("engineering", "Engineering", "cto")
                    .Position("platform", "Nina Patel", "Platform Lead", "cto", "engineering")
                    .Position("security", "Owen Brooks", "Security Lead", "cto", "engineering")
                    .Callout(" cto ", "cto-note", "Owns platform and security roadmap", 8.1, 5.9, options => {
                        options.Width = 2.65;
                        options.Height = 0.72;
                        options.RouteOffset = 0.1;
                    }));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape callout = Assert.Single(page.Callouts());
            VisioShape target = Assert.Single(page.Shapes, shape => shape.Id == "cto");
            Assert.Equal("cto-note", callout.Id);
            Assert.Equal("Owns platform and security roadmap", callout.Text);
            Assert.Equal(target.Id, callout.CalloutTargetId);
            Assert.Contains("Annotations", callout.LayerNames);
            Assert.Equal(2.65, callout.Width);
            Assert.Equal(0.72, callout.Height);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, callout));
            Assert.Same(target, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);
            Assert.Contains("Annotations", leader.LayerNames);
            Assert.Equal(leader.Id, callout.GetUserCellValue("OfficeIMO.CalloutLeaderId"));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void OrgChartDiagramBuilderCanAutoPlaceSemanticCalloutsBesideNodes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .OrgChartDiagram("Auto Annotated Leadership", org => org
                    .Title()
                    .Root("ceo", "Marta Nowak", "Chief Executive Officer")
                    .Manager("cto", "Alex Chen", "Chief Technology Officer", "ceo")
                    .TeamBand("engineering", "Engineering", "cto")
                    .Position("platform", "Nina Patel", "Platform Lead", "cto", "engineering")
                    .Position("security", "Owen Brooks", "Security Lead", "cto", "engineering")
                    .Callout("cto", "cto-note", "Owns platform and security roadmap", VisioSide.Right, 0.45, options => {
                        options.Width = 2.65;
                        options.Height = 0.72;
                    })
                    .Callout("platform", "Succession backup", VisioSide.Bottom, 0.25));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape cto = Assert.Single(page.Shapes, shape => shape.Id == "cto");
            VisioShape platform = Assert.Single(page.Shapes, shape => shape.Id == "platform");
            VisioShape explicitCallout = Assert.Single(page.Callouts(), shape => shape.Id == "cto-note");
            VisioShape generatedCallout = Assert.Single(page.Callouts(), shape => shape.Id == "platform-callout");

            Assert.True(explicitCallout.PinX > cto.PinX);
            Assert.Equal(cto.PinY, explicitCallout.PinY, 6);
            Assert.Equal(cto.Id, explicitCallout.CalloutTargetId);
            Assert.Equal(2.65, explicitCallout.Width);
            Assert.True(generatedCallout.PinY < platform.PinY);
            Assert.Equal(platform.Id, generatedCallout.CalloutTargetId);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, explicitCallout));
            Assert.Same(cto, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void OrgChartDiagramBuilderGeneratesUniqueCalloutIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .OrgChartDiagram("Generated", org => org
                    .Root("ceo", "CEO")
                    .Manager("cto", "CTO", "Technology", "ceo")
                    .Callout("cto", "First note", 6.3, 5.8)
                    .Callout("cto", "Second note", 6.3, 4.9));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(new[] { "cto-callout", "cto-callout-2" }, page.Callouts().Select(shape => shape.Id).ToArray());
        }

        [Fact]
        public void OrgChartDiagramBuilderRejectsUnknownManager() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.OrgChartDiagram("Invalid", org => org
                    .Root("ceo", "CEO")
                    .Position("missing", "Person", "Role", "unknown")));

            Assert.Contains("Unknown org chart node id", exception.Message);
        }

        [Fact]
        public void OrgChartDiagramBuilderRejectsGeneratedShapeIdCollisions() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException nodeFirst = Assert.Throws<ArgumentException>(() =>
                document.OrgChartDiagram("Invalid", org => org
                    .Root("title", "CEO")
                    .Title()));
            ArgumentException titleFirst = Assert.Throws<ArgumentException>(() =>
                document.OrgChartDiagram("Invalid", org => org
                    .Title()
                    .Root("title", "CEO")));
            ArgumentException bandShapeCollision = Assert.Throws<ArgumentException>(() =>
                document.OrgChartDiagram("Invalid", org => org
                    .Root("ceo", "CEO")
                    .Position("org-band-engineering", "Dev", "Lead", "ceo")
                    .TeamBand("engineering", "Engineering", "ceo")));

            Assert.Contains("already exists", nodeFirst.Message);
            Assert.Contains("already exists", titleFirst.Message);
            Assert.Contains("already exists", bandShapeCollision.Message);
        }

        [Fact]
        public void OrgChartDiagramBuilderRejectsCalloutIdCollisionsAndUnknownTargets() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknownTarget = Assert.Throws<ArgumentException>(() =>
                document.OrgChartDiagram("Invalid", org => org
                    .Root("ceo", "CEO")
                    .Callout("missing", "note", "No target", 4, 4)));
            ArgumentException nodeCollision = Assert.Throws<ArgumentException>(() =>
                document.OrgChartDiagram("Invalid", org => org
                    .Root("ceo", "CEO")
                    .Callout("ceo", "ceo", "Duplicate id", 4, 4)));
            ArgumentException bandCollision = Assert.Throws<ArgumentException>(() =>
                document.OrgChartDiagram("Invalid", org => org
                    .Root("ceo", "CEO")
                    .TeamBand("leadership", "Leadership", "ceo")
                    .Callout("ceo", "org-band-leadership", "Duplicate id", 4, 4)));

            Assert.Contains("Unknown org chart node id", unknownTarget.Message);
            Assert.Contains("already exists", nodeCollision.Message);
            Assert.Contains("already exists", bandCollision.Message);
        }

        [Fact]
        public void OrgChartDiagramBuilderRejectsAutoCalloutPlacementIssues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentOutOfRangeException autoPlacement = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.OrgChartDiagram("Invalid", org => org
                    .Root("ceo", "CEO")
                    .Callout("ceo", "Invalid", VisioSide.Auto)));
            ArgumentOutOfRangeException badGap = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.OrgChartDiagram("Invalid", org => org
                    .Root("ceo", "CEO")
                    .Callout("ceo", "Invalid", VisioSide.Right, double.NaN)));

            Assert.Contains("Placement must be", autoPlacement.Message);
            Assert.Contains("finite non-negative", badGap.Message);
        }
    }
}
