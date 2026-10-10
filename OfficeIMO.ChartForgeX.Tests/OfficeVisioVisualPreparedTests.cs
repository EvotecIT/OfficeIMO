using System;
using System.IO;
using System.Linq;
using ChartForgeX.Primitives;
using ChartForgeX.Rendering;
using ChartForgeX.Topology;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed partial class OfficeVisioVisualIntegrationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PreparedTopologyRetainsResolvedRoutesLabelsAndViewportThroughSave(bool jsonIngress) {
        var chart = TopologyChart.Create().WithViewport(640, 400);
        chart.LayoutMode = TopologyLayoutMode.Manual;
        chart.Nodes.Add(new TopologyNode { Id = "api", Label = "API", X = 40, Y = 100, Width = 120, Height = 64 });
        chart.Nodes.Add(new TopologyNode { Id = "worker", Label = "Worker", X = 380, Y = 240, Width = 120, Height = 64 });
        chart.Edges.Add(new TopologyEdge { Id = "request", SourceNodeId = "api", TargetNodeId = "worker", Label = "Request", Routing = TopologyEdgeRouting.ObstacleAvoidingOrthogonal });
        var prepared = chart.Prepare(new VisualRenderContext(
            layout: new VisualLayoutOptions(new VisualSize(640, 400)), frame: new VisualFrame(showLegend: false)));
        var artifact = prepared.ToArtifact("prepared-topology", VisualArtifactKind.Topology);
        var envelope = artifact.ToInterchangeEnvelope();
        var edge = Assert.Single(envelope.Edges);
        Assert.Empty(edge.Topology!.Waypoints);
        Assert.True(edge.ResolvedRoute.Count >= 2);
        Assert.NotNull(edge.ResolvedLabelBounds);
        var options = new OfficeVisioVisualOptions { PixelsPerInch = 144, LayoutMode = OfficeVisioVisualLayoutMode.Preserve };
        var result = jsonIngress ? artifact.ToInterchangeUtf8Json().ToOfficeVisio(options) : artifact.ToOfficeVisio(options);
        Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.LayoutRecomputed ||
            item.Feature == "computedRoute" || item.Feature == "selfRoute");

        void Check(VisioPage page) {
            double ppi = options.PixelsPerInch;
            Assert.Equal(envelope.Width!.Value / ppi, page.Width, 6);
            Assert.Equal(envelope.Height!.Value / ppi, page.Height, 6);
            foreach (var source in envelope.Nodes) {
                var shape = page.Shapes.Single(item => item.Id == source.Id);
                Assert.Equal((source.X!.Value + source.Width!.Value / 2) / ppi, shape.PinX, 6);
                Assert.Equal((envelope.Height.Value - source.Y!.Value - source.Height!.Value / 2) / ppi, shape.PinY, 6);
                Assert.Equal(source.Width.Value / ppi, shape.Width, 6);
                Assert.Equal(source.Height.Value / ppi, shape.Height, 6);
            }
            var connector = Assert.Single(page.Connectors);
            Assert.NotNull(connector.From);
            Assert.NotNull(connector.To);
            Assert.Equal(VisioConnectorRerouteBehavior.Never, connector.RerouteBehavior);
            Assert.Equal(edge.ResolvedRoute.Count - 2, connector.Waypoints.Count);
            var first = edge.ResolvedRoute[0];
            Assert.Equal(first.X / ppi - connector.From.PinX + connector.From.Width / 2, connector.FromConnectionPoint!.X, 6);
            Assert.Equal((envelope.Height.Value - first.Y) / ppi - connector.From.PinY + connector.From.Height / 2, connector.FromConnectionPoint.Y, 6);
            var last = edge.ResolvedRoute[edge.ResolvedRoute.Count - 1];
            Assert.Equal(last.X / ppi - connector.To.PinX + connector.To.Width / 2, connector.ToConnectionPoint!.X, 6);
            Assert.Equal((envelope.Height.Value - last.Y) / ppi - connector.To.PinY + connector.To.Height / 2, connector.ToConnectionPoint.Y, 6);
            for (int index = 0; index < connector.Waypoints.Count; index++) {
                Assert.Equal(edge.ResolvedRoute[index + 1].X / ppi, connector.Waypoints[index].X, 6);
                Assert.Equal((envelope.Height.Value - edge.ResolvedRoute[index + 1].Y) / ppi, connector.Waypoints[index].Y, 6);
            }
            var label = edge.ResolvedLabelBounds!.Value;
            var bounds = connector.GetLabelBounds();
            Assert.Equal(label.X / ppi, bounds.Left, 6);
            Assert.Equal((envelope.Height.Value - label.Y - label.Height) / ppi, bounds.Bottom, 6);
            Assert.Equal((label.X + label.Width) / ppi, bounds.Right, 6);
            Assert.Equal((envelope.Height.Value - label.Y) / ppi, bounds.Top, 6);
        }

        Check(result.Page);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            result.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            Check(VisioDocument.Load(path).Pages[0]);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PreparedFlowAndSequenceRemainEditableAndReportNativeReflowThroughSave(bool sequence) {
        PreparedVisual prepared;
        if (sequence) {
            var source = SequenceArtifact.Create("sequence").WithSize(640, 400)
                .AddParticipant("api", "API").AddParticipant("worker", "Worker").AddMessage("api", "worker", "Request");
            prepared = source.Prepare(new VisualRenderContext(
                layout: new VisualLayoutOptions(new VisualSize(640, 400)), frame: new VisualFrame(showLegend: false)));
        } else {
            var source = FlowArtifact.Create("flow").AddLane("team", "Team")
                .AddStep("api", "API", FlowArtifactStepKind.Start, "team")
                .AddStep("worker", "Worker", FlowArtifactStepKind.Decision, "team").AddConnector("api", "worker", "Request");
            prepared = source.Prepare(new VisualRenderContext(
                layout: new VisualLayoutOptions(new VisualSize(640, 400)), frame: new VisualFrame(showLegend: false)));
        }
        var artifact = prepared.ToArtifact(sequence ? "prepared-sequence" : "prepared-flow",
            sequence ? VisualArtifactKind.Sequence : VisualArtifactKind.Flow);
        var result = artifact.ToInterchangeUtf8Json().ToOfficeVisio(new OfficeVisioVisualOptions { UseNaturalPageSize = true });
        Assert.True(result.Report.AllProjectedObjectsEditable);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
        Assert.Throws<NotSupportedException>(() => artifact.ToOfficeVisio(new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Preserve }));
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            result.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            var page = VisioDocument.Load(path).Pages[0];
            Assert.Contains(page.Shapes, shape => (sequence ? shape.GetShapeDataValue("CFX.Id") : shape.Id) == "api");
            Assert.Contains(page.Shapes, shape => (sequence ? shape.GetShapeDataValue("CFX.Id") : shape.Id) == "worker");
            Assert.Contains(page.Connectors, connector => connector.Label == "Request");
        } finally { if (File.Exists(path)) File.Delete(path); }
    }
}
