using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioArchitectureDiagramBuilderTests {
        [Fact]
        public void ArchitectureDiagramBuilderCreatesStyledInfrastructurePage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .ArchitectureDiagram("Jenkins on Azure", diagram => diagram
                    .Theme(VisioStyleTheme.Technical())
                    .Region("vnet", "Virtual Network", 1, 0, 4, 3)
                    .Region("subnet", "Build Subnet", 1, 1, 4, 2)
                    .Actor("users", "Users", 0, 1)
                    .Gateway("public-ip", "Public IP", 1, 1)
                    .Service("jenkins", "Jenkins Server", 2, 1)
                    .Compute("agent", "Build Agent", 3, 1)
                    .Database("data", "Data", 2, 2)
                    .Storage("artifacts", "Artifacts", 4, 2)
                    .Security("vault", "Key Vault", 2, 0)
                    .DataFlow("users", "public-ip", "HTTPS")
                    .DataFlow("public-ip", "jenkins", "route")
                    .ControlFlow("jenkins", "agent", "scale")
                    .Dependency("jenkins", "data", "state")
                    .Dependency("jenkins", "vault", "secrets")
                    .DataFlow("agent", "artifacts", "publish"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Jenkins on Azure", page.Name);
            Assert.Equal(9, page.Shapes.Count);
            Assert.Equal(6, page.Connectors.Count);
            Assert.Contains(page.Shapes, shape => shape.Id == "vnet" && shape.NameU == "Rectangle");
            Assert.Contains(page.Shapes, shape => shape.Id == "jenkins" && shape.NameU == "Process");
            Assert.Contains(page.Shapes, shape => shape.Id == "data" && shape.NameU == "Data");
            Assert.Contains(page.Shapes, shape => shape.Id == "vault" && shape.NameU == "Decision");
            Assert.All(page.Connectors, connector => Assert.NotEmpty(connector.Waypoints));
            Assert.All(page.Connectors.Where(connector => !string.IsNullOrWhiteSpace(connector.Label)), connector => Assert.NotNull(connector.LabelPlacement));
            Assert.Empty(page.AnalyzeVisualQuality().Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(9, loaded.Pages[0].Shapes.Count);
            Assert.Equal(6, loaded.Pages[0].Connectors.Count);
        }

        [Fact]
        public void ArchitectureStencilCatalogExposesCommonInfrastructureShapes() {
            Assert.Equal("Architecture", VisioStencils.Architecture.Name);
            Assert.Equal("Compute", VisioStencils.Architecture.Get("vm").Name);
            Assert.Equal("Security", VisioStencils.Architecture.Get("identity").Name);
            Assert.Equal("Queue", VisioStencils.All.Get("arch.queue").Name);
        }

        [Fact]
        public void ArchitectureDiagramBuilderRejectsUnknownLinkEndpoints() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.ArchitectureDiagram("Invalid", diagram => diagram
                    .Service("api", "API", 0, 0)
                    .DataFlow("api", "missing")));

            Assert.Contains("Unknown architecture component id", exception.Message);
        }
    }
}
