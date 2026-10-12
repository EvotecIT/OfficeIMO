using System.Collections.Generic;
using System.Linq;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;

namespace OfficeIMO.ChartForgeX;

public static partial class OfficeVisioVisualConversionExtensions {
    private static void ApplyResolvedLabelGeometry(VisioPage page,
        VisualArtifactInterchangeEnvelope envelope, OfficeVisioVisualOptions options) {
        var edges = envelope.Edges.Where(edge => edge.ResolvedLabelBounds.HasValue)
            .ToDictionary(edge => edge.Id, edge => edge.ResolvedLabelBounds!.Value);
        foreach (var connector in page.Connectors) {
            if (string.IsNullOrWhiteSpace(connector.Label) || !edges.TryGetValue(connector.Id, out var bounds)) continue;
            double ppi = options.PixelsPerInch;
            connector.LabelPlacement = VisioConnectorLabelPlacement.At(
                (bounds.X + bounds.Width / 2) / ppi,
                (envelope.Height!.Value - bounds.Y - bounds.Height / 2) / ppi,
                bounds.Width / ppi, bounds.Height / ppi);
        }
    }

    private static void ReportPreservedRouteBounds(IReadOnlyList<VisioConnectorWaypoint> points,
        VisualArtifactInterchangeEnvelope envelope, double ppi,
        OfficeVisioVisualConversionReport report, string edgeId) {
        if (points.Any(point => point.X < 0 || point.Y < 0 || point.X > envelope.Width!.Value / ppi || point.Y > envelope.Height!.Value / ppi)) {
            report.Warn(OfficeVisioVisualDiagnosticCode.GeometryOutsidePage, OfficeVisioVisualEntityKind.Edge, edgeId, "route",
                "The preserved connector extends beyond the declared page.");
        }
    }
}
