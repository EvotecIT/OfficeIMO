using global::ChartForgeX.Core;
using global::ChartForgeX.Primitives;
using global::ChartForgeX.Rendering;
using global::ChartForgeX.Themes;
using global::ChartForgeX.Topology;
using global::ChartForgeX.VisualArtifacts;

namespace OfficeIMO.ChartForgeX.Examples;

internal sealed record VisualSpecimen(string Name, VisualThemeMode Mode, VisualArtifact Artifact, bool EditableTopology) {
    public string ModeName => Mode.ToString().ToLowerInvariant();
    public string Caption => Artifact.Accessibility.Description ?? Artifact.Title;
    public OfficeVisualConversionResult Convert(double usableWidthPoints) => Artifact.ToOfficeVisual(
        new OfficeVisualConversionOptions {
            WidthPoints = Math.Min(450D, usableWidthPoints),
            SvgPolicy = OfficeVisualSvgPolicy.RasterizeWhenNeeded
        });
}

internal static class VisualSpecimens {
    public const double LogicalWidth = 600D;
    public const double LogicalHeight = 360D;

    public static IReadOnlyList<VisualSpecimen> Create() {
        var result = new List<VisualSpecimen>();
        foreach (var mode in new[] { VisualThemeMode.Light, VisualThemeMode.Dark }) {
            var chart = Chart.Create()
                .AddLine("Observed", new[] { new ChartPoint(1, 42), new ChartPoint(2, 58), new ChartPoint(3, 73),
                    new ChartPoint(4, 91), new ChartPoint(5, 67), new ChartPoint(6, 82) })
                .AddLine("Capacity", new[] { new ChartPoint(1, 100), new ChartPoint(2, 100), new ChartPoint(3, 100),
                    new ChartPoint(4, 100), new ChartPoint(5, 100), new ChartPoint(6, 100) });
            const string loadDescription = "Observed events per minute across six intervals: 42, 58, 73, 91, 67 and 82. Capacity is 100 in every interval.";
            chart.WithAccessibility(value => value.WithTextAlternative("Service capacity", loadDescription, "en"));
            var chartArtifact = chart.Prepare(Context(mode, "Service capacity", "Events per minute · six observation intervals", true))
                .ToArtifact("service-capacity");
            result.Add(new VisualSpecimen("capacity-" + mode.ToString().ToLowerInvariant(), mode, chartArtifact, false));

            var topology = TopologyChart.Create().WithViewport(LogicalWidth, LogicalHeight);
            topology.LayoutMode = TopologyLayoutMode.Manual;
            topology.Nodes.Add(new TopologyNode { Id = "gateway", Label = "Gateway", X = 32, Y = 150, Width = 148, Height = 72 });
            topology.Nodes.Add(new TopologyNode { Id = "worker", Label = "Processing worker", X = 202, Y = 150, Width = 176, Height = 72 });
            topology.Nodes.Add(new TopologyNode { Id = "ledger", Label = "Ledger store", X = 420, Y = 150, Width = 148, Height = 72 });
            topology.Edges.Add(new TopologyEdge { Id = "accept", SourceNodeId = "gateway", TargetNodeId = "worker", Label = "Accept",
                Routing = TopologyEdgeRouting.ObstacleAvoidingOrthogonal });
            topology.Edges.Add(new TopologyEdge { Id = "persist", SourceNodeId = "worker", TargetNodeId = "ledger", Label = "Persist",
                Routing = TopologyEdgeRouting.ObstacleAvoidingOrthogonal });
            var topologyArtifact = topology.Prepare(Context(mode, "Service delivery", "Three editable services · prepared routes", false))
                .ToArtifact("service-delivery");
            topologyArtifact.Accessibility.WithTextAlternative("Service delivery",
                "Gateway accepts requests into the Processing worker. The worker persists them in the Ledger store.", "en");
            result.Add(new VisualSpecimen("delivery-" + mode.ToString().ToLowerInvariant(), mode, topologyArtifact, true));
        }
        // Reuse the exact artifact after the dark variant to exercise document media identity independently of its Id.
        result.Add(result[0] with { Name = "capacity-light-repeat" });
        return result;
    }

    private static VisualRenderContext Context(VisualThemeMode mode, string title, string subtitle, bool legend) =>
        OfficeVisualDocumentStyle.Default.CreateContext(themeMode: mode,
            frame: new VisualFrame(title, subtitle, showLegend: legend, legendPosition: ChartLegendPosition.BottomLeft,
                showSurface: true));
}
