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
    }
}
