using global::ChartForgeX.Core;
using global::ChartForgeX.Primitives;
using global::ChartForgeX.Rendering;
using global::ChartForgeX.Topology;
using global::ChartForgeX.VisualArtifacts;
using global::ChartForgeX.VisualBlocks;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

// One test per row of the supported-families table in the adapter README: each family is produced by its
// ChartForgeX builder, saved, validated and reopened before shapes, connectors, Shape Data and diagnostics are checked.
public sealed class OfficeVisioVisualSupportedFamilyTests {
    private static readonly VisualRenderContext Context = new(
        layout: new VisualLayoutOptions(new VisualSize(640, 400)), frame: new VisualFrame("Delivery", showLegend: false));

    [Fact]
    public void TopologyKeepsPreparedShapesConnectorsShapeDataAndReportsAdornmentLoss() {
        var chart = TopologyChart.Create().WithViewport(640, 400);
        chart.LayoutMode = TopologyLayoutMode.Manual;
        chart.Nodes.Add(new TopologyNode { Id = "api", Label = "API", Kind = TopologyNodeKind.Service, IconId = "server",
            Href = "https://example.test/api", X = 40, Y = 120, Width = 140, Height = 64 });
        chart.Nodes.Add(new TopologyNode { Id = "db", Label = "Orders", Kind = TopologyNodeKind.Database, X = 420, Y = 120, Width = 140, Height = 64 });
        chart.Edges.Add(new TopologyEdge { Id = "write", SourceNodeId = "api", TargetNodeId = "db", Label = "Write",
            Routing = TopologyEdgeRouting.ObstacleAvoidingOrthogonal });
        VisualArtifact artifact = chart.Prepare(Context).ToArtifact("topology", VisualArtifactKind.Topology);

        OfficeVisioVisualConversionResult result = artifact.ToOfficeVisio(
            new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Preserve });

        Assert.Equal(VisualArtifactInterchangeFamily.Topology, result.Report.SemanticFamily);
        Assert.Equal(OfficeVisioVisualProjectionKind.Graph, result.Report.Projection);
        Assert.DoesNotContain(result.Report.NativeDiagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
        AssertLoss(result, OfficeVisioVisualDiagnosticCode.ArtworkNotProjected, "api");
        VisioPage page = SaveAndReload(result);
        VisioShape api = page.Shapes.Single(shape => shape.Id == "api");
        Assert.Equal("Service", api.GetShapeDataValue("CFX.TopologyNodeKind"));
        Assert.Equal("server", api.GetShapeDataValue("CFX.Icon"));
        Assert.Contains(api.Hyperlinks, link => link.Address == "https://example.test/api");
        Assert.Equal("Database", page.Shapes.Single(shape => shape.Id == "db").GetShapeDataValue("CFX.TopologyNodeKind"));
        VisioConnector write = page.Connectors.Single(connector => connector.Id == "write");
        Assert.Equal("api", write.From!.Id);
        Assert.Equal("db", write.To!.Id);
        Assert.Equal("Write", write.Label);
    }

    [Fact]
    public void FlowKeepsStepsConnectorsShapeDataAndReportsNativeReflow() {
        VisualArtifact artifact = FlowArtifact.Create("flow").AddLane("ops", "Operations")
            .AddStep("submit", "Submit", FlowArtifactStepKind.Start, "ops")
            .AddStep("review", "Review", FlowArtifactStepKind.Decision, "ops", VisualStatus.Warning)
            .AddStep("deploy", "Deploy", FlowArtifactStepKind.Process, "ops")
            .AddConnector("submit", "review", "Request").AddConnector("review", "deploy", "Approved")
            .Prepare(Context).ToArtifact("flow", VisualArtifactKind.Flow);

        OfficeVisioVisualConversionResult result = artifact.ToOfficeVisio();

        Assert.Equal(VisualArtifactInterchangeFamily.Flow, result.Report.SemanticFamily);
        Assert.Equal(OfficeVisioVisualProjectionKind.FlowGraph, result.Report.Projection);
        AssertLoss(result, OfficeVisioVisualDiagnosticCode.LayoutRecomputed, "flow");
        Assert.Throws<NotSupportedException>(() => artifact.ToOfficeVisio(
            new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Preserve }));
        VisioPage page = SaveAndReload(result);
        Assert.Equal("Start", page.Shapes.Single(shape => shape.Id == "submit").GetShapeDataValue("CFX.FlowStepKind"));
        VisioShape review = page.Shapes.Single(shape => shape.Id == "review");
        Assert.Equal("Decision", review.GetShapeDataValue("CFX.FlowStepKind"));
        Assert.Equal("Warning", review.GetShapeDataValue("CFX.Status"));
        VisioConnector approved = page.Connectors.Single(connector => connector.Label == "Approved");
        Assert.Equal("review", approved.From!.Id);
        Assert.Equal("deploy", approved.To!.Id);
        Assert.Equal(2, page.Connectors.Count);
    }

    [Fact]
    public void SequenceKeepsParticipantsMessagesShapeDataAndReportsNativeReflow() {
        VisualArtifact artifact = SequenceArtifact.Create("sequence").WithSize(640, 400)
            .AddParticipant("client", "Client").AddParticipant("api", "API")
            .AddMessage("client", "api", "Create order").AddMessage("api", "client", "Created")
            .Prepare(Context).ToArtifact("sequence", VisualArtifactKind.Sequence);

        OfficeVisioVisualConversionResult result = artifact.ToOfficeVisio();

        Assert.Equal(VisualArtifactInterchangeFamily.Sequence, result.Report.SemanticFamily);
        Assert.Equal(OfficeVisioVisualProjectionKind.Sequence, result.Report.Projection);
        AssertLoss(result, OfficeVisioVisualDiagnosticCode.LayoutRecomputed, "sequence");
        VisioPage page = SaveAndReload(result);
        Assert.Contains(page.Shapes, shape => shape.GetShapeDataValue("CFX.Id") == "client" &&
            shape.GetShapeDataValue("CFX.SequenceParticipantOrder") == "0");
        Assert.Contains(page.Shapes, shape => shape.GetShapeDataValue("CFX.Id") == "api");
        VisioConnector create = page.Connectors.Single(connector => connector.Label == "Create order");
        Assert.NotNull(create.GetShapeDataValue("CFX.SequenceMessageKind"));
        Assert.Contains(page.Connectors, connector => connector.Label == "Created");
    }

    [Fact]
    public void ChartsAndBlocksFailClosedForVisioAndKeepTheirPicturePlacement() {
        VisualArtifact chart = Chart.Create().AddLine("Load", new[] { new ChartPoint(1, 2), new ChartPoint(2, 3) })
            .Prepare(Context).ToArtifact("chart");
        VisualArtifact table = TableArtifact.Create("table").AddColumn("name", "Name").AddRow("api", "API").ToVisualArtifact();

        foreach (VisualArtifact artifact in new[] { chart, table }) {
            Assert.Throws<NotSupportedException>(() => artifact.ToOfficeVisio());
            OfficeVisualConversionResult picture = artifact.ToOfficeVisual();
            Assert.NotEmpty(picture.GetPlacementBytes());
        }
    }

    private static void AssertLoss(OfficeVisioVisualConversionResult result, OfficeVisioVisualDiagnosticCode code, string entityId) {
        Assert.True(result.Report.AllProjectedObjectsEditable);
        Assert.True(result.Report.HasSemanticLoss);
        Assert.Contains(result.Report.NativeDiagnostics, item => item.Code == code && item.EntityId == entityId &&
            item.Severity == OfficeVisioVisualDiagnosticSeverity.Warning);
        Assert.Throws<OfficeVisioVisualFidelityException>(() => result.Envelope.ToOfficeVisio(
            new OfficeVisioVisualOptions { RequireLossless = true }));
    }

    private static VisioPage SaveAndReload(OfficeVisioVisualConversionResult result) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".vsdx");
        try {
            result.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            return Assert.Single(VisioDocument.Load(path).Pages);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
