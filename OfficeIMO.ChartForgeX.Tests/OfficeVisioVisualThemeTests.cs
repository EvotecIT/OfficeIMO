using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ChartForgeX.Primitives;
using ChartForgeX.Rendering;
using ChartForgeX.Themes;
using ChartForgeX.Topology;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed partial class OfficeVisioVisualIntegrationTests {
    [Theory]
    [InlineData(VisualThemeMode.Light, true)]
    [InlineData(VisualThemeMode.Dark, true)]
    [InlineData(VisualThemeMode.Dark, false)]
    public void PreparedThemePageAndLabelFillsPersistBehindEditableTopology(VisualThemeMode mode, bool preparedRoute) {
        var chart = TopologyChart.Create().WithViewport(640, 400);
        chart.LayoutMode = TopologyLayoutMode.Manual;
        chart.Nodes.Add(new TopologyNode { Id = "api", Label = "API", X = 40, Y = 160, Width = 120, Height = 64 });
        chart.Nodes.Add(new TopologyNode { Id = "worker", Label = "Worker", X = 380, Y = 160, Width = 120, Height = 64 });
        chart.Edges.Add(new TopologyEdge { Id = "request", SourceNodeId = "api", TargetNodeId = "worker", Label = "Request",
            Routing = TopologyEdgeRouting.ObstacleAvoidingOrthogonal });
        var theme = VisualTheme.Graphite();
        var artifact = chart.Prepare(new VisualRenderContext(
            layout: new VisualLayoutOptions(new VisualSize(640, 400)), theme: theme, themeMode: mode,
            frame: new VisualFrame("Delivery", "Prepared editable services", showLegend: false)))
            .ToArtifact("themed-topology", VisualArtifactKind.Topology);
        var envelope = artifact.ToInterchangeEnvelope();
        if (!preparedRoute) {
            // The adapter's fallback routing includes title adornments, but must exclude the canvas fill.
            envelope.Edges[0].ResolvedRoute.Clear();
            envelope.Edges[0].ResolvedLabelBounds = null;
            envelope.Edges[0].Topology!.Waypoints.Clear();
        }
        string expectedFill = theme.Resolve(mode).Background.ToHex();
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 100,
            LayoutMode = OfficeVisioVisualLayoutMode.Preserve });
        Assert.True(result.Report.AllProjectedObjectsEditable);
        Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.LayoutRecomputed);

        void Check(VisioPage page) {
            var fill = page.Shapes[0];
            Assert.Equal("OfficeIMO Page Background", fill.NameU);
            Assert.True(fill.IsBackgroundSurface);
            Assert.Equal(expectedFill, fill.FillColor.ToString(), ignoreCase: true);
            Assert.Equal(page.Width, fill.Width, 6);
            Assert.Equal(page.Height, fill.Height, 6);
            Assert.Equal(page.Width / 2, fill.PinX, 6);
            Assert.Equal(page.Height / 2, fill.PinY, 6);
            Assert.Equal(0, fill.LinePattern);
            Assert.True(fill.Protection.LockSelect);
            foreach (var source in envelope.Nodes) {
                var shape = page.Shapes.Single(item => item.Id == source.Id);
                Assert.Equal(source.Label, shape.Text);
                Assert.Equal((source.X!.Value + source.Width!.Value / 2) / 100, shape.PinX, 6);
                Assert.Equal(source.Width.Value / 100, shape.Width, 6);
                Assert.NotEqual(true, shape.Protection.LockSelect);
            }
            var connector = Assert.Single(page.Connectors);
            Assert.Equal("request", connector.Id);
            Assert.Equal("Request", connector.Label);
            Assert.Equal(expectedFill, connector.TextStyle!.BackgroundColor!.Value.ToString(), ignoreCase: true);
            Assert.Equal(VisioConnectorRerouteBehavior.Never, connector.RerouteBehavior);
            Assert.All(connector.Waypoints, point => {
                Assert.InRange(point.X, 0D, page.Width);
                Assert.InRange(point.Y, 0D, page.Height);
            });
            XNamespace ns = "http://www.w3.org/2000/svg";
            var svg = XDocument.Parse(page.ToSvg());
            var canvas = svg.Descendants(ns + "g").Single(item => item.Attribute("data-officeimo-visio-page") != null);
            Assert.Equal(fill.Id, (string?)canvas.Elements(ns + "g").First().Attribute("data-visio-shape-id"));
            var background = canvas.Elements(ns + "g").First().Elements(ns + "path").First();
            Assert.Equal(expectedFill, (string?)background.Attribute("fill"), ignoreCase: true);
            var caption = svg.Descendants(ns + "rect").Single(item => item.Attribute("data-officeimo-connector-label-background") != null);
            Assert.Equal(expectedFill, (string?)caption.Attribute("fill"), ignoreCase: true);
        }

        Check(result.Page);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            result.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            var loaded = VisioDocument.Load(path);
            Check(Assert.Single(loaded.Pages));
        } finally { if (File.Exists(path)) File.Delete(path); }
    }
}
