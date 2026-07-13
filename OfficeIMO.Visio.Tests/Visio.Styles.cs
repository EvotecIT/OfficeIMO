using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Fluent;
using OfficeIMO.Visio.Stencils;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioStyleTests {
        [Fact]
        public void ShapeAndConnectorStylesApplyToModelsSelectionsAndFluentApi() {
            VisioStyleTheme theme = VisioStyleTheme.Minimal();
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Styles");
            VisioShape first = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "first", 2, 4, "First");
            VisioShape second = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "second", 5, 4, "Second");
            first.Data["Style"] = "Primary";
            second.Data["Style"] = "Primary";
            VisioConnector connector = page.AddConnector(first, second, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);

            first.ApplyStyle(theme.Decision);
            page.SelectWithData("Style", "Primary").Style(theme.Primary);
            connector.ApplyStyle(theme.Connector);
            page.SelectConnectedConnectors(first).Style(theme.ControlConnector);

            Assert.Equal(theme.Primary.FillColor, first.FillColor);
            Assert.Equal(theme.Primary.LineColor, second.LineColor);
            Assert.Equal(theme.ControlConnector.LineColor, connector.LineColor);
            Assert.Equal(theme.ControlConnector.LinePattern, connector.LinePattern);
            Assert.Equal(theme.ControlConnector.EndArrow, connector.EndArrow);
            Assert.Equal(theme.Primary.TextStyle!.Color, first.TextStyle!.Color);
            Assert.Equal(theme.ControlConnector.TextStyle!.Color, connector.TextStyle!.Color);

            VisioDocument fluentDocument = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            fluentDocument.AsFluent()
                .Page("Fluent", pageBuilder => pageBuilder
                    .Rect("a", 1, 1, 2, 1, "A")
                    .Shape("a", shape => shape.Style(theme.Success))
                    .Rect("b", 4, 1, 2, 1, "B")
                    .Connect("a", "b", VisioSide.Right, VisioSide.Left, connection => connection.Style(theme.DataConnector)))
                .End();

            Assert.Equal(theme.Success.FillColor, fluentDocument.Pages[0].Shapes[0].FillColor);
            Assert.Equal(theme.DataConnector.LineColor, fluentDocument.Pages[0].Connectors[0].LineColor);
            Assert.Equal(theme.Success.TextStyle!.Color, fluentDocument.Pages[0].Shapes[0].TextStyle!.Color);
            Assert.Equal(theme.DataConnector.TextStyle!.Color, fluentDocument.Pages[0].Connectors[0].TextStyle!.Color);
        }

        [Fact]
        public void StyleThemeCloneIsDetached() {
            VisioStyleTheme original = VisioStyleTheme.Technical();
            VisioStyleTheme clone = original.Clone();

            clone.Primary.FillColor = Color.Red;
            clone.Primary.TextStyle!.Color = Color.Red;
            clone.Connector.LineColor = Color.Green;
            clone.Connector.TextStyle!.Color = Color.Green;

            Assert.NotEqual(original.Primary.FillColor, clone.Primary.FillColor);
            Assert.NotEqual(original.Primary.TextStyle!.Color, clone.Primary.TextStyle!.Color);
            Assert.NotEqual(original.Connector.LineColor, clone.Connector.LineColor);
            Assert.NotEqual(original.Connector.TextStyle!.Color, clone.Connector.TextStyle!.Color);
            Assert.Equal("OfficeIMO Technical", original.Name);
        }

        [Fact]
        public void AdditionalThemePresetsCarryReadableTextStyles() {
            VisioStyleTheme office = VisioStyleTheme.Office();
            VisioStyleTheme fluent = VisioStyleTheme.Fluent();
            VisioStyleTheme technical = VisioStyleTheme.Technical();
            VisioStyleTheme dark = VisioStyleTheme.Dark();
            VisioStyleTheme enterprise = VisioStyleTheme.Enterprise();
            VisioStyleTheme cloud = VisioStyleTheme.Cloud();
            VisioStyleTheme process = VisioStyleTheme.Process();
            VisioStyleTheme print = VisioStyleTheme.Print();
            VisioStyleTheme darkSafe = VisioStyleTheme.DarkSafe();

            Assert.Equal("OfficeIMO Office", office.Name);
            Assert.Equal("OfficeIMO Fluent", fluent.Name);
            Assert.Equal("OfficeIMO Technical", technical.Name);
            Assert.Equal("OfficeIMO Dark", dark.Name);
            Assert.Equal("OfficeIMO Enterprise", enterprise.Name);
            Assert.Equal("OfficeIMO Cloud", cloud.Name);
            Assert.Equal("OfficeIMO Process", process.Name);
            Assert.Equal("OfficeIMO Print", print.Name);
            Assert.Equal("OfficeIMO Dark Safe", darkSafe.Name);
            Assert.NotNull(office.Primary.TextStyle);
            Assert.NotNull(fluent.Decision.TextStyle);
            Assert.NotNull(technical.ControlConnector.TextStyle);
            Assert.NotNull(dark.Primary.TextStyle);
            Assert.NotNull(enterprise.Container.TextStyle);
            Assert.NotNull(cloud.DataConnector.TextStyle);
            Assert.NotNull(process.ControlConnector.TextStyle);
            Assert.NotNull(print.Decision.TextStyle);
            Assert.NotNull(darkSafe.Container.TextStyle);
            Assert.Equal(Color.White, dark.Primary.TextStyle!.Color);
            Assert.Equal(Color.White, dark.Success.TextStyle!.Color);
            Assert.NotEqual(dark.Container.FillColor, dark.Container.TextStyle!.Color);
            Assert.NotNull(dark.Connector.TextStyle);
            Assert.NotEqual(technical.Primary.FillColor, technical.Marker.FillColor);
            Assert.NotEqual(enterprise.Primary.FillColor, enterprise.Decision.FillColor);
            Assert.NotEqual(cloud.Primary.FillColor, cloud.Container.FillColor);
            Assert.NotEqual(process.Primary.FillColor, process.Emphasis.FillColor);
            Assert.Equal(Color.Black, print.Primary.TextStyle!.Color);
            Assert.Equal(2, print.Marker.LinePattern);
            Assert.Equal(2, print.ControlConnector.LinePattern);
            Assert.NotEqual(darkSafe.Container.FillColor, darkSafe.Container.TextStyle!.Color);
            Assert.True(darkSafe.Primary.LineWeight > VisioShape.DefaultLineWeight);
        }

        [Fact]
        public void DiagramBuildersUseOneStyleThemeAndLayoutOnlyOptions() {
            VisioStyleTheme enterprise = VisioStyleTheme.Enterprise();
            VisioStyleTheme technical = VisioStyleTheme.Technical();
            var flowLayout = new VisioFlowchartLayoutOptions { ProcessWidth = 3.1D };
            var blockLayout = new VisioBlockDiagramLayoutOptions { ColumnGap = 1.4D };

            Assert.All(typeof(VisioFlowchartLayoutOptions).GetProperties(), property => Assert.Equal(typeof(double), property.PropertyType));
            Assert.All(typeof(VisioBlockDiagramLayoutOptions).GetProperties(), property => Assert.Equal(typeof(double), property.PropertyType));
            Assert.Equal(3.1D, flowLayout.Clone().ProcessWidth);
            Assert.Equal(1.4D, blockLayout.Clone().ColumnGap);
            Assert.NotEqual(enterprise.Primary.FillColor, enterprise.Decision.FillColor);
            Assert.NotEqual(technical.DataConnector.LinePattern, technical.ControlConnector.LinePattern);
            Assert.NotNull(enterprise.TitleText);
            Assert.NotNull(technical.LegendText);
        }

        [Fact]
        public void PremiumPresetSetIncludesTechnicalPrintAndDarkSafeFamilies() {
            VisioStyleTheme[] presets = VisioStyleTheme.PremiumPresets().ToArray();

            Assert.Equal(new[] {
                "OfficeIMO Enterprise",
                "OfficeIMO Technical",
                "OfficeIMO Cloud",
                "OfficeIMO Process",
                "OfficeIMO Print",
                "OfficeIMO Dark Safe"
            }, presets.Select(theme => theme.Name).ToArray());
            Assert.All(presets, theme => Assert.NotNull(theme.Primary.TextStyle));
            Assert.All(presets, theme => Assert.NotNull(theme.Connector.TextStyle));
            Assert.All(presets, theme => Assert.True(theme.Primary.LineWeight >= 0.016));
        }

        [Fact]
        public void DiagramTitleStylePreservesThemeColor() {
            VisioStyleTheme theme = VisioStyleTheme.Technical();
            theme.Emphasis.TextStyle!.Color = Color.Red;

            VisioTextStyle style = VisioDiagramTitleStyles.Create(theme);

            Assert.Equal(Color.Red, style.Color);
        }

        [Fact]
        public void TechnicalAndPrintThemesGenerateValidatedDiagrams() {
            string technicalPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string printPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(technicalPath)
                .BlockDiagram("Technical System", diagram => diagram
                    .Theme(VisioStyleTheme.Technical())
                    .Region("platform", "Platform", 0, 0, 3, 1)
                    .Block("api", "API", 0, 0)
                    .EmphasisBlock("worker", "Worker", 1, 0)
                    .Block("queue", "Queue", 2, 0)
                    .DataFlow("api", "worker", "request")
                    .ControlFlow("worker", "queue", "dispatch"))
                .Save();

            VisioDocument.Create(printPath)
                .Flowchart("Print Safe Approval", flow => flow
                    .Theme(VisioStyleTheme.Print())
                    .Start("start", "Request")
                    .Step("review", "Review")
                    .Decision("approved", "Approved?")
                    .End("done", "Done"))
                .Save();

            Assert.Empty(VisioValidator.Validate(technicalPath));
            Assert.Empty(VisioValidator.Validate(printPath));
        }

        [Fact]
        public void StyleThemeCanDriveFlowchartBuilder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioStyleTheme theme = VisioStyleTheme.Minimal();

            VisioDocument document = VisioDocument.Create(filePath)
                .Flowchart("Styled Flowchart", flow => flow
                    .Theme(theme)
                    .Start("start", "Start")
                    .Step("step", "Do work")
                    .Decision("decision", "Done?")
                    .End("end", "End"));

            VisioPage page = document.Pages[0];

            Assert.Equal(theme.Success.FillColor, page.Shapes.Single(shape => shape.Id == "start").FillColor);
            Assert.Equal(theme.Primary.FillColor, page.Shapes.Single(shape => shape.Id == "step").FillColor);
            Assert.Equal(theme.Decision.FillColor, page.Shapes.Single(shape => shape.Id == "decision").FillColor);
            Assert.All(page.Connectors, connector => Assert.Equal(theme.Connector.LineColor, connector.LineColor));
            Assert.All(page.Shapes, shape => Assert.NotNull(shape.TextStyle));
            Assert.All(page.Connectors, connector => Assert.NotNull(connector.TextStyle));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void StyleThemeCanDriveBlockDiagramBuilder() {
            VisioStyleTheme theme = VisioStyleTheme.Dark();

            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .BlockDiagram("Styled Blocks", diagram => diagram
                    .Theme(theme)
                    .Region("region", "Zone", 0, 0, 2, 1)
                    .Block("input", "Input", 0, 0)
                    .EmphasisBlock("worker", "Worker", 1, 0)
                    .ControlFlow("input", "worker", "control"));

            VisioPage page = document.Pages[0];
            VisioShape region = page.Shapes.Single(shape => shape.Id == "region");
            VisioShape input = page.Shapes.Single(shape => shape.Id == "input");
            VisioShape worker = page.Shapes.Single(shape => shape.Id == "worker");
            VisioConnector connector = page.Connectors.Single();

            Assert.Equal(theme.Container.FillColor, region.FillColor);
            Assert.Equal(theme.Primary.LineColor, input.LineColor);
            Assert.Equal(theme.Emphasis.FillColor, worker.FillColor);
            Assert.Equal(theme.ControlConnector.LineColor, connector.LineColor);
            Assert.Equal(theme.ControlConnector.LinePattern, connector.LinePattern);
            Assert.Equal(theme.Container.TextStyle!.Color, region.TextStyle!.Color);
            Assert.Equal(theme.Primary.TextStyle!.Color, input.TextStyle!.Color);
            Assert.Equal(theme.Emphasis.TextStyle!.Color, worker.TextStyle!.Color);
            Assert.Equal(theme.ControlConnector.TextStyle!.Color, connector.TextStyle!.Color);
        }
    }
}
