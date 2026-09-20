using System;
using System.Collections.Generic;
using System.Linq;
using global::ChartForgeX.Topology;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.ChartForgeX;

public static partial class OfficeVisioVisualConversionExtensions {
    private static bool ShouldPreserveLayout(VisualArtifactInterchangeEnvelope envelope, OfficeVisioVisualOptions options) {
        if (options.LayoutMode == OfficeVisioVisualLayoutMode.Reflow) return false;
        bool complete = envelope.Family == VisualArtifactInterchangeFamily.Topology &&
            envelope.Width > 0 && envelope.Height > 0 &&
            envelope.Nodes.All(node => HasBounds(node.X, node.Y, node.Width, node.Height)) &&
            (!options.IncludeGroups || envelope.Groups.All(group => HasBounds(group.X, group.Y, group.Width, group.Height)));
        if (!complete && options.LayoutMode == OfficeVisioVisualLayoutMode.Preserve) {
            throw new NotSupportedException("Preserving geometry requires a topology envelope with complete node, group, and viewport bounds. Use Reflow explicitly for other inputs.");
        }
        return complete;
    }

    private static bool HasBounds(double? x, double? y, double? width, double? height) =>
        x.HasValue && y.HasValue && width > 0 && height > 0;

    private static void ApplyPreparedGeometry(
        VisualArtifactInterchangeEnvelope envelope, OfficeVisioVisualOptions options,
        List<VisioGraphNodeRecord> nodes, List<VisioGraphEdgeRecord> edges,
        List<VisioGraphClusterRecord> groups, OfficeVisioVisualConversionReport report) {
        double ppi = options.PixelsPerInch, pageHeight = envelope.Height!.Value;
        var sourceNodes = envelope.Nodes.ToDictionary(node => node.Id, StringComparer.Ordinal);
        var sourceGroups = envelope.Groups.ToDictionary(group => group.Id, StringComparer.Ordinal);
        var sourceEdges = envelope.Edges.ToDictionary(edge => edge.Id, StringComparer.Ordinal);
        foreach (var node in nodes) {
            var source = sourceNodes[node.Id];
            ReportBounds(source.X!.Value, source.Y!.Value, source.Width!.Value, source.Height!.Value, envelope, report, OfficeVisioVisualEntityKind.Node, source.Id);
            node.Placement = Placement(source.X!.Value, source.Y!.Value, source.Width!.Value, source.Height!.Value, ppi, pageHeight);
        }
        foreach (var group in groups) {
            var source = sourceGroups[group.Id];
            ReportBounds(source.X!.Value, source.Y!.Value, source.Width!.Value, source.Height!.Value, envelope, report, OfficeVisioVisualEntityKind.Group, source.Id);
            group.Placement = Placement(source.X!.Value, source.Y!.Value, source.Width!.Value, source.Height!.Value, ppi, pageHeight);
        }
        foreach (var edge in edges) {
            var source = sourceEdges[edge.Id!];
            var from = sourceNodes[source.SourceId];
            var to = sourceNodes[source.TargetId];
            var waypoints = source.Topology!.Waypoints;
            var start = Endpoint(from, source.SourcePortId, source.Topology.SourcePort,
                waypoints.Count > 0 ? waypoints[0].X : to.X!.Value + to.Width!.Value / 2,
                waypoints.Count > 0 ? waypoints[0].Y : to.Y!.Value + to.Height!.Value / 2);
            var end = Endpoint(to, source.TargetPortId, source.Topology.TargetPort,
                waypoints.Count > 0 ? waypoints[waypoints.Count - 1].X : from.X!.Value + from.Width!.Value / 2,
                waypoints.Count > 0 ? waypoints[waypoints.Count - 1].Y : from.Y!.Value + from.Height!.Value / 2);
            var points = new List<VisioConnectorWaypoint> { Point(start.X, start.Y, ppi, pageHeight) };
            foreach (var point in waypoints) points.Add(Point(point.X, point.Y, ppi, pageHeight));
            if (waypoints.Count == 0 && source.SourceId == source.TargetId) {
                double right = from.X!.Value + from.Width!.Value, top = from.Y!.Value;
                if (string.IsNullOrEmpty(source.SourcePortId) && source.Topology.SourcePort == TopologyEdgePort.Auto)
                    start = (right, top + from.Height!.Value * 0.25);
                if (string.IsNullOrEmpty(source.TargetPortId) && source.Topology.TargetPort == TopologyEdgePort.Auto)
                    end = (right, top + from.Height!.Value * 0.75);
                points.Clear();
                points.Add(Point(start.X, start.Y, ppi, pageHeight));
                points.Add(Point(right + 32, start.Y, ppi, pageHeight));
                if (Math.Abs(start.Y - end.Y) < 1e-9) {
                    points.Add(Point(right + 32, top - 32, ppi, pageHeight));
                    points.Add(Point(end.X, top - 32, ppi, pageHeight));
                } else points.Add(Point(right + 32, end.Y, ppi, pageHeight));
                report.Warn(OfficeVisioVisualDiagnosticCode.EdgePresentationNormalized, OfficeVisioVisualEntityKind.Edge, source.Id, "selfRoute",
                    "The self relationship was projected as a native loop; its computed CFX route was not present in the envelope.");
            } else if (waypoints.Count == 0 && source.Topology.Routing != TopologyEdgeRouting.Straight) {
                // The envelope contains authored waypoints, not the renderer's computed obstacle route.
                // Keep the bounds and attachments, but report this visible route normalization.
                points.Add(Point((start.X + end.X) / 2, start.Y, ppi, pageHeight));
                points.Add(Point((start.X + end.X) / 2, end.Y, ppi, pageHeight));
                report.Warn(OfficeVisioVisualDiagnosticCode.EdgePresentationNormalized, OfficeVisioVisualEntityKind.Edge, source.Id, "computedRoute",
                    "Node bounds and attachments were preserved, but the envelope did not contain a resolved route; a native orthogonal route was used.");
            }
            points.Add(Point(end.X, end.Y, ppi, pageHeight));
            if (points.Any(point => point.X < 0 || point.Y < 0 || point.X > envelope.Width!.Value / ppi || point.Y > pageHeight / ppi))
                report.Warn(OfficeVisioVisualDiagnosticCode.GeometryOutsidePage, OfficeVisioVisualEntityKind.Edge, source.Id, "route",
                    "The preserved connector extends beyond the declared page.");
            edge.Route = new VisioGraphRoute(points);
        }
    }

    private static void ConfigurePreservedTitle(VisioGraphDiagramBuilder builder, VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualOptions options, IEnumerable<VisioGraphEdgeRecord> edges, OfficeVisioVisualConversionReport report) {
        if (!options.IncludeTitle || !HasTitle(envelope)) return;
        double available = envelope.Nodes.Select(node => node.Y!.Value)
            .Concat(options.IncludeGroups ? envelope.Groups.Select(group => group.Y!.Value) : Array.Empty<double>())
            .Concat(edges.Where(edge => edge.Route != null).SelectMany(edge => edge.Route!.Points)
                .Select(point => envelope.Height!.Value - point.Y * options.PixelsPerInch))
            .DefaultIfEmpty(envelope.Height!.Value).Min() / options.PixelsPerInch;
        const double margin = 0.16, height = 0.45, gap = 0.08;
        if (available < margin + height + gap) {
            report.Warn(OfficeVisioVisualDiagnosticCode.TitleNotProjected, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "title",
                "The preserved geometry leaves no clear header band for a native title. The title remains in the source envelope and document metadata.");
            return;
        }
        builder.Margins(0.4, margin, 0.4, 0.4).Title(CombineLabel(envelope.Title, envelope.Subtitle), UniqueTitleId(envelope), height, gap);
    }

    private static void RouteComputedConnectors(VisioPage page, VisualArtifactInterchangeEnvelope envelope, OfficeVisioVisualConversionReport report) {
        var computed = new HashSet<string>(envelope.Edges.Where(edge => edge.Topology!.Waypoints.Count == 0 &&
            edge.SourceId != edge.TargetId && edge.Topology.Routing != TopologyEdgeRouting.Straight).Select(edge => edge.Id), StringComparer.Ordinal);
        foreach (var connector in page.Connectors.Where(connector => computed.Contains(connector.Id))) {
            connector.RouteOrthogonalAroundShapes(page.Shapes, new VisioConnectorRoutingOptions { IncludeDiagramAdornments = true });
            if (connector.Waypoints.Any(point => point.X < 0 || point.Y < 0 || point.X > page.Width || point.Y > page.Height))
                report.Warn(OfficeVisioVisualDiagnosticCode.GeometryOutsidePage, OfficeVisioVisualEntityKind.Edge, connector.Id, "route",
                    "The normalized obstacle route extends beyond the declared page.");
        }
    }

    private static void ReportBounds(double x, double y, double width, double height,
        VisualArtifactInterchangeEnvelope envelope, OfficeVisioVisualConversionReport report,
        OfficeVisioVisualEntityKind kind, string id) {
        if (x < 0 || y < 0 || x + width > envelope.Width || y + height > envelope.Height)
            report.Warn(OfficeVisioVisualDiagnosticCode.GeometryOutsidePage, kind, id, "bounds",
                "Preserved bounds extend beyond the declared page. Increase the source viewport or use Reflow.");
    }

    private static VisioGraphPlacement Placement(double x, double y, double width, double height, double ppi, double pageHeight) =>
        new((x + width / 2) / ppi, (pageHeight - y - height / 2) / ppi, width / ppi, height / ppi);

    private static VisioConnectorWaypoint Point(double x, double y, double ppi, double pageHeight) => new(x / ppi, (pageHeight - y) / ppi);

    private static (double X, double Y) Endpoint(VisualArtifactInterchangeNode node, string? portId, TopologyEdgePort side, double towardX, double towardY) {
        double x = node.X!.Value, y = node.Y!.Value, width = node.Width!.Value, height = node.Height!.Value, offset = 0.5;
        if (!string.IsNullOrWhiteSpace(portId)) {
            var port = node.Ports.Single(item => item.Id == portId);
            side = port.Side; offset = port.Offset;
        }
        return side switch {
            TopologyEdgePort.Top => (x + width * offset, y),
            TopologyEdgePort.Bottom => (x + width * offset, y + height),
            TopologyEdgePort.Left => (x, y + height * offset),
            TopologyEdgePort.Right => (x + width, y + height * offset),
            _ => RectangleEndpoint(x, y, width, height, towardX, towardY)
        };
    }

    private static (double X, double Y) RectangleEndpoint(double x, double y, double width, double height, double towardX, double towardY) {
        double cx = x + width / 2, cy = y + height / 2, dx = towardX - cx, dy = towardY - cy;
        if (Math.Abs(dx) + Math.Abs(dy) < 1e-9) return (x + width, cy);
        double scale = 1 / Math.Max(Math.Abs(dx) / (width / 2), Math.Abs(dy) / (height / 2));
        return (cx + dx * scale, cy + dy * scale);
    }

    private static void EnforceFidelity(OfficeVisioVisualConversionReport report, OfficeVisioVisualOptions options) {
        if (report.Diagnostics.Any(item => options.RejectedDiagnostics.Contains(item.Code) ||
            options.RequireLossless && item.Severity == OfficeVisioVisualDiagnosticSeverity.Warning)) {
            throw new OfficeVisioVisualFidelityException(report);
        }
    }
}

/// <summary>A native projection did not satisfy the caller's fidelity policy. No document was saved.</summary>
public sealed class OfficeVisioVisualFidelityException : InvalidOperationException {
    internal OfficeVisioVisualFidelityException(OfficeVisioVisualConversionReport report) : base("Native Visio conversion did not satisfy the requested fidelity policy. Inspect Report.Diagnostics for details.") { Report = report; }
    /// <summary>Gets all diagnostics collected before the policy rejected the result.</summary>
    public OfficeVisioVisualConversionReport Report { get; }
}
