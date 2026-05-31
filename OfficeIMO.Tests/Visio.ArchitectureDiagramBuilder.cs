using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

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
            Assert.Equal(11, page.Shapes.Count);
            Assert.Equal(6, page.Connectors.Count);
            VisioShape vnet = Assert.Single(page.Shapes, shape => shape.Id == "vnet" && shape.NameU == "Rectangle");
            VisioShape vnetLabel = Assert.Single(page.Shapes, shape => shape.Id == "vnet-label" && shape.Text == "Virtual Network");
            Assert.True(vnet.IsBackgroundSurface);
            Assert.Equal(string.Empty, vnet.Text);
            Assert.Equal("Text Box", vnetLabel.NameU);
            Assert.True(vnetLabel.PinY > vnet.PinY + vnet.Height / 2D);
            Assert.Contains(page.Shapes, shape => shape.Id == "jenkins" && shape.NameU == "Process");
            Assert.Contains(page.Shapes, shape => shape.Id == "data" && shape.NameU == "Data");
            Assert.Contains(page.Shapes, shape => shape.Id == "vault" && shape.NameU == "Decision");
            VisioStencilProfile profile = document.CreateStencilProfile();
            Assert.Equal(7, profile.StencilBackedShapeCount);
            Assert.Equal(new[] { "Architecture" }, profile.StencilCatalogs);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "arch.service" && usage.Count == 1);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "arch.compute" && usage.Count == 1);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "arch.security" && usage.Count == 1);
            Assert.All(page.Connectors, connector => Assert.NotEmpty(connector.Waypoints));
            Assert.All(page.Connectors.Where(connector => !string.IsNullOrWhiteSpace(connector.Label)), connector => Assert.NotNull(connector.LabelPlacement));
            string[] qualityIssues = page.AnalyzeVisualQuality().Select(issue => issue.ToString()).ToArray();
            Assert.True(qualityIssues.Length == 0, string.Join(Environment.NewLine, qualityIssues));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(11, loaded.Pages[0].Shapes.Count);
            Assert.Equal(6, loaded.Pages[0].Connectors.Count);
        }

        [Fact]
        public void ArchitectureStencilCatalogExposesCommonInfrastructureShapes() {
            Assert.Equal("Architecture", VisioStencils.Architecture.Name);
            Assert.Equal("Compute", VisioStencils.Architecture.Get("vm").Name);
            Assert.Equal("Security", VisioStencils.Architecture.Get("identity").Name);
            Assert.Equal("Queue", VisioStencils.All.Get("arch.queue").Name);
            Assert.Equal("External System", VisioStencils.Architecture.Get("third-party").Name);
        }

        [Fact]
        public void ArchitectureDiagramBuilderCanAddTitleWithoutOverlappingComponents() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .ArchitectureDiagram("Jenkins on Azure", diagram => diagram
                    .Title()
                    .Region("vnet", "Virtual Network", 1, 0, 3, 2)
                    .Actor("users", "Users", 0, 1)
                    .Gateway("public-ip", "Public IP", 1, 1)
                    .Service("jenkins", "Jenkins Server", 2, 1)
                    .Compute("agent", "Build Agent", 3, 1)
                    .DataFlow("users", "public-ip", "HTTPS")
                    .DataFlow("public-ip", "jenkins")
                    .ControlFlow("jenkins", "agent", "scale"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape title = Assert.Single(page.Shapes, shape => shape.Id == "title");
            VisioShape gateway = Assert.Single(page.Shapes, shape => shape.Id == "public-ip");
            Assert.Equal("Text Box", title.NameU);
            Assert.Equal("Jenkins on Azure", title.Text);
            Assert.True(title.PinY > gateway.PinY);
            Assert.Empty(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabels = false
            }).Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void ArchitectureDiagramBuilderCanAddLegendWithoutOverlappingComponents() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .ArchitectureDiagram("Jenkins on Azure", diagram => diagram
                    .Title()
                    .Legend(dataFlowLabel: "Data Flow", controlFlowLabel: "Control Flow")
                    .Region("vnet", "Virtual Network", 1, 0, 3, 2)
                    .Actor("users", "Users", 0, 1)
                    .Gateway("public-ip", "Public IP", 1, 1)
                    .Service("jenkins", "Jenkins Server", 2, 1)
                    .Compute("agent", "Build Agent", 3, 1)
                    .DataFlow("users", "public-ip", "HTTPS")
                    .DataFlow("public-ip", "jenkins")
                    .ControlFlow("jenkins", "agent", "scale"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape dataLegend = Assert.Single(page.Shapes, shape => shape.Text == "Data Flow");
            VisioShape controlLegend = Assert.Single(page.Shapes, shape => shape.Text == "Control Flow");
            VisioShape gateway = Assert.Single(page.Shapes, shape => shape.Id == "public-ip");
            Assert.Equal("Text Box", dataLegend.NameU);
            Assert.Equal("Text Box", controlLegend.NameU);
            Assert.True(dataLegend.PinY > gateway.PinY);
            Assert.True(controlLegend.PinY > gateway.PinY);
            Assert.Contains(page.Shapes, shape => shape.NameU == "Rectangle" && shape.LinePattern == 1 && shape.Text.Length == 0);
            Assert.Contains(page.Shapes, shape => shape.NameU == "Rectangle" && shape.LinePattern == 2 && shape.Text.Length == 0);
            Assert.Empty(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabels = false
            }).Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void ArchitectureDiagramBuilderCanAddSemanticCallouts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .ArchitectureDiagram("Jenkins on Azure", diagram => diagram
                    .Theme(VisioStyleTheme.Technical())
                    .Service("jenkins", "Jenkins Server", 1, 1)
                    .Compute("agent", "Build Agent", 2, 1)
                    .ControlFlow("jenkins", "agent", "scale")
                    .Callout("jenkins", "scale-note", "Scale agents on demand", 5.8, 6.6, options => {
                        options.Width = 2.45;
                        options.Height = 0.72;
                        options.ShapeStyle = new VisioShapeStyle(Color.FromRgb(239, 246, 252), Color.FromRgb(0, 120, 212), 0.014);
                        options.RouteOffset = 0.12;
                    }));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape callout = Assert.Single(page.Callouts());
            VisioShape target = Assert.Single(page.Shapes, shape => shape.Id == "jenkins");
            Assert.Equal("scale-note", callout.Id);
            Assert.Equal("Scale agents on demand", callout.Text);
            Assert.Equal(target.Id, callout.CalloutTargetId);
            Assert.Contains("Annotations", callout.LayerNames);
            Assert.Equal(2.45, callout.Width);
            Assert.Equal(0.72, callout.Height);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, callout));
            Assert.Same(target, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);
            Assert.Contains("Annotations", leader.LayerNames);
            Assert.Equal(leader.Id, callout.GetUserCellValue("OfficeIMO.CalloutLeaderId"));
            Assert.Empty(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabels = false
            }).Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void ArchitectureDiagramBuilderCanAutoPlaceSemanticCalloutsBesideComponents() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .ArchitectureDiagram("Auto Annotated Architecture", diagram => diagram
                    .Title()
                    .Region("vnet", "Virtual Network", 0, 0, 3, 2)
                    .Service("jenkins", "Jenkins", 1, 0)
                    .Compute("agent", "Build Agent", 2, 0)
                    .Security("vault", "Key Vault", 1, 1)
                    .ControlFlow("jenkins", "agent", "scale")
                    .Dependency("jenkins", "vault", "secrets")
                    .Callout("jenkins", "scale-note", "Scale agents automatically", VisioSide.Right, 0.45, options => {
                        options.Width = 2.5;
                        options.Height = 0.72;
                    })
                    .Callout("vault", "Managed identity boundary", VisioSide.Bottom, 0.25));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape jenkins = Assert.Single(page.Shapes, shape => shape.Id == "jenkins");
            VisioShape vault = Assert.Single(page.Shapes, shape => shape.Id == "vault");
            VisioShape explicitCallout = Assert.Single(page.Callouts(), shape => shape.Id == "scale-note");
            VisioShape generatedCallout = Assert.Single(page.Callouts(), shape => shape.Id == "vault-callout");

            Assert.True(explicitCallout.PinX > jenkins.PinX);
            Assert.Equal(jenkins.PinY, explicitCallout.PinY, 6);
            Assert.Equal(jenkins.Id, explicitCallout.CalloutTargetId);
            Assert.Equal(2.5, explicitCallout.Width);
            Assert.True(generatedCallout.PinY < vault.PinY);
            Assert.Equal(vault.Id, generatedCallout.CalloutTargetId);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, explicitCallout));
            Assert.Same(jenkins, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
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

        [Fact]
        public void ArchitectureDiagramBuilderRejectsUnknownCalloutTargets() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.ArchitectureDiagram("Invalid", diagram => diagram
                    .Service("api", "API", 0, 0)
                    .Callout("missing", "note", "No target", 3, 3)));

            Assert.Contains("Unknown architecture component id", exception.Message);
        }

        [Fact]
        public void ArchitectureDiagramBuilderRejectsTitleIdCollisions() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.ArchitectureDiagram("Invalid", diagram => diagram
                    .Title(id: "api")
                    .Service("api", "API", 0, 0)));

            Assert.Contains("already exists", exception.Message);
        }

        [Fact]
        public void ArchitectureDiagramBuilderRejectsCalloutIdCollisions() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.ArchitectureDiagram("Invalid", diagram => diagram
                    .Service("api", "API", 0, 0)
                    .Callout("api", "api", "Duplicate id", 3, 3)));

            Assert.Contains("already exists", exception.Message);
        }

        [Fact]
        public void ArchitectureDiagramBuilderRejectsAutoCalloutPlacementIssues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentOutOfRangeException autoPlacement = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.ArchitectureDiagram("Invalid", diagram => diagram
                    .Service("api", "API", 0, 0)
                    .Callout("api", "Invalid", VisioSide.Auto)));
            ArgumentOutOfRangeException badGap = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.ArchitectureDiagram("Invalid", diagram => diagram
                    .Service("api", "API", 0, 0)
                    .Callout("api", "Invalid", VisioSide.Right, double.NaN)));

            Assert.Contains("Placement must be", autoPlacement.Message);
            Assert.Contains("finite non-negative", badGap.Message);
        }
    }
}
