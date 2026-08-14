using System;
using global::ChartForgeX.Topology;
using global::ChartForgeX.VisualArtifacts;

namespace OfficeIMO.ChartForgeX;

public static partial class OfficeVisioVisualConversionExtensions {
    private static void ReportArtifactAccessibilityFidelity(
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualConversionReport report) {
        if (!string.IsNullOrWhiteSpace(envelope.AccessibleName) ||
            !string.IsNullOrWhiteSpace(envelope.AccessibleDescription) ||
            !string.IsNullOrWhiteSpace(envelope.Language) ||
            envelope.IsDecorative) {
            report.Warn(OfficeVisioVisualDiagnosticCode.AccessibilityNotProjected, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "accessibility",
                "Artifact accessibility and language semantics remain in the CFX envelope because the native Visio projection has no equivalent page-level contract.");
        }
    }

    private static void ReportArtifactPresentationFidelity(
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualConversionReport report) {
        if (envelope.Presentation?.Legend != null) {
            report.Warn(OfficeVisioVisualDiagnosticCode.PresentationNotProjected, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "legend",
                "The resolved CFX legend remains in the semantic envelope because the native Visio projection does not create a legend block.");
        }
        if (envelope.Presentation?.MapViewport != null) {
            report.Warn(OfficeVisioVisualDiagnosticCode.PresentationNotProjected, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "mapViewport",
                "The geographic viewport remains in the semantic envelope because native Visio graph layout does not preserve map projection semantics.");
        }
        if (envelope.Presentation?.Theme != null) {
            report.Warn(OfficeVisioVisualDiagnosticCode.PresentationNotProjected, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "theme",
                "Resolved entity colors were projected where supported; the complete CFX theme remains in the semantic envelope.");
        }
    }

    private static void ReportTitleFidelity(
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        if (HasTitle(envelope) && !options.IncludeTitle) {
            report.Warn(OfficeVisioVisualDiagnosticCode.TitleNotProjected, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "title",
                "The visible CFX title was omitted from the editable Visio page by conversion options and remains available as document metadata.");
        }
    }

    private static void ReportScenarioFidelity(
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualConversionReport report) {
        foreach (VisualArtifactInterchangeScenario scenario in envelope.Scenarios) {
            report.Warn(OfficeVisioVisualDiagnosticCode.ScenarioNotProjected, OfficeVisioVisualEntityKind.Artifact, scenario.Id, "scenario",
                $"Scenario '{scenario.Id}' remains in the CFX envelope because native Visio diagrams do not project guided scenario playback.");
        }
    }

    private static void ReportGraphSemanticFidelity(
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualConversionReport report,
        bool flow) {
        foreach (VisualArtifactInterchangeNode node in envelope.Nodes) {
            if (node.Details.Count > 0) {
                report.Warn(OfficeVisioVisualDiagnosticCode.DetailsNotRendered, OfficeVisioVisualEntityKind.Node, node.Id, "details",
                    $"Node '{node.Id}' detail rows remain in the CFX envelope and, when enabled, Shape Data because the editable native Visio shape does not render them.");
            }
            if (!string.IsNullOrWhiteSpace(node.IconId) || !string.IsNullOrWhiteSpace(node.Symbol) || !string.IsNullOrWhiteSpace(node.Badge)) {
                report.Warn(OfficeVisioVisualDiagnosticCode.ArtworkNotProjected, OfficeVisioVisualEntityKind.Node, node.Id, "nodeAdornment",
                    $"Node '{node.Id}' icon, symbol, or badge remains in the CFX envelope because the editable native Visio graph shape does not render those adornments.");
            }
            if (flow) continue;
            if (node.Topology!.Artwork != null) {
                report.Warn(OfficeVisioVisualDiagnosticCode.ArtworkNotProjected, OfficeVisioVisualEntityKind.Node, node.Id, "artwork",
                    $"Node '{node.Id}' portable artwork remains in the CFX envelope because the native graph projection selected an editable Visio stencil.");
            }
            if (node.Topology.DisplayMode is not TopologyNodeDisplayMode.Card and not TopologyNodeDisplayMode.CompactCard) {
                report.Warn(OfficeVisioVisualDiagnosticCode.SemanticLoss, OfficeVisioVisualEntityKind.Node, node.Id, "displayMode",
                    $"Node '{node.Id}' display mode '{node.Topology.DisplayMode}' was normalized to an editable native Visio graph shape.");
            }
        }
        if (flow) return;
        foreach (VisualArtifactInterchangeGroup group in envelope.Groups) {
            if (!string.IsNullOrWhiteSpace(group.Topology!.IconId) || !string.IsNullOrWhiteSpace(group.Topology.Symbol)) {
                report.Warn(OfficeVisioVisualDiagnosticCode.ArtworkNotProjected, OfficeVisioVisualEntityKind.Group, group.Id, "headerArtwork",
                    $"Group '{group.Id}' header artwork remains in the CFX envelope because native Visio containers do not project that header mark.");
            }
        }
        foreach (VisualArtifactInterchangeEdge edge in envelope.Edges) {
            VisualArtifactInterchangeTopologyEdge topology = edge.Topology!;
            if (topology.Waypoints.Count > 0 || topology.DashPattern.Count > 0 || topology.SourceMarker.HasValue || topology.TargetMarker.HasValue ||
                topology.StrokeWidth.HasValue || topology.Opacity.HasValue || topology.IsMuted || topology.RoutingPriority != 0 || topology.RouteLane.HasValue ||
                topology.LabelOffsetX != 0D || topology.LabelOffsetY != 0D || topology.LabelAnchor != null || topology.LabelAnchorNodeId != null ||
                topology.LayoutInference != TopologyEdgeLayoutInference.None || topology.PreferredLength.HasValue || topology.MinimumRankSpan != 0 ||
                topology.Routing is TopologyEdgeRouting.Curved or TopologyEdgeRouting.ObstacleAvoidingOrthogonal ||
                topology.Emphasis != TopologyEdgeEmphasis.Normal) {
                report.Warn(OfficeVisioVisualDiagnosticCode.EdgePresentationNormalized, OfficeVisioVisualEntityKind.Edge, edge.Id, "edgePresentation",
                    $"Edge '{edge.Id}' advanced routing or presentation remains in the CFX envelope because native Visio graph layout recomputed the connector.");
            }
        }
    }

    private static void ReportSequenceSemanticFidelity(
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualConversionReport report) {
        foreach (VisualArtifactInterchangeNode participant in envelope.Nodes) {
            if (!string.IsNullOrWhiteSpace(participant.IconId) || !string.IsNullOrWhiteSpace(participant.Symbol) || !string.IsNullOrWhiteSpace(participant.Badge)) {
                report.Warn(OfficeVisioVisualDiagnosticCode.ArtworkNotProjected, OfficeVisioVisualEntityKind.Participant, participant.Id, "participantAdornment",
                    $"Sequence participant '{participant.Id}' icon, symbol, or badge remains in the CFX envelope and, when enabled, Shape Data because the native Visio participant does not render those adornments.");
            }
            if (participant.Details.Count > 0) {
                report.Warn(OfficeVisioVisualDiagnosticCode.DetailsNotRendered, OfficeVisioVisualEntityKind.Participant, participant.Id, "details",
                    $"Sequence participant '{participant.Id}' detail rows remain in the CFX envelope and, when enabled, Shape Data because the native Visio participant does not render them.");
            }
        }
        foreach (VisualArtifactInterchangeEdge message in envelope.Edges) {
            if (!string.IsNullOrWhiteSpace(message.SourceLabel) || !string.IsNullOrWhiteSpace(message.TargetLabel)) {
                report.Warn(OfficeVisioVisualDiagnosticCode.EndpointLabelsNotRendered, OfficeVisioVisualEntityKind.Message, message.Id, "endpointLabels",
                    $"Sequence message '{message.Id}' endpoint labels remain in the CFX envelope and, when enabled, Shape Data because the native Visio message does not render endpoint labels.");
            }
        }
    }
}
