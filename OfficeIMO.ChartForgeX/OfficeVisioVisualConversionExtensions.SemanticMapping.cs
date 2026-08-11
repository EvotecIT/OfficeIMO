using System;
using global::ChartForgeX.Primitives;
using global::ChartForgeX.Topology;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.ChartForgeX;

public static partial class OfficeVisioVisualConversionExtensions {
    private static VisioGraphLayout MapLayout(VisualArtifactInterchangeEnvelope envelope, bool flow, OfficeVisioVisualConversionReport report) {
        if (flow) {
            if (envelope.Flow!.LayoutMode == FlowArtifactLayoutMode.Force) {
                report.Warn(OfficeVisioVisualDiagnosticCode.LayoutNormalized, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "layoutMode",
                    "CFX force-directed flow layout was normalized to Visio's native radial layout.");
                return VisioGraphLayout.Radial;
            }
            return envelope.Flow.LayoutMode == FlowArtifactLayoutMode.Dense ? VisioGraphLayout.Grid : VisioGraphLayout.Layered;
        }
        switch (envelope.Topology!.LayoutMode) {
            case TopologyLayoutMode.GroupGrid:
            case TopologyLayoutMode.Matrix:
            case TopologyLayoutMode.DenseGrouped:
                return VisioGraphLayout.Grid;
            case TopologyLayoutMode.HubAndSpoke:
            case TopologyLayoutMode.RelationshipRadial:
            case TopologyLayoutMode.MindMap:
                return VisioGraphLayout.Radial;
            case TopologyLayoutMode.ForceDirected:
                report.Warn(OfficeVisioVisualDiagnosticCode.LayoutNormalized, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "layoutMode",
                    "CFX force-directed topology layout was normalized to Visio's native radial layout.");
                return VisioGraphLayout.Radial;
            case TopologyLayoutMode.Geographic:
                report.Warn(OfficeVisioVisualDiagnosticCode.LayoutNormalized, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "layoutMode",
                    "CFX geographic layout was normalized to Visio's native layered layout; coordinates remain in the envelope.");
                return VisioGraphLayout.Layered;
            default:
                return VisioGraphLayout.Layered;
        }
    }

    private static VisioGraphDirection MapDirection(VisualArtifactInterchangeEnvelope envelope, bool flow, OfficeVisioVisualConversionReport report) {
        bool vertical;
        bool reverse;
        string source;
        if (flow) {
            FlowArtifactDirection direction = envelope.Flow!.LayoutDirection;
            vertical = direction is FlowArtifactDirection.TopToBottom or FlowArtifactDirection.BottomToTop;
            reverse = direction is FlowArtifactDirection.BottomToTop or FlowArtifactDirection.RightToLeft;
            source = direction.ToString();
        } else {
            TopologyLayoutDirection direction = envelope.Topology!.LayoutDirection;
            vertical = direction is TopologyLayoutDirection.TopToBottom or TopologyLayoutDirection.BottomToTop;
            reverse = direction is TopologyLayoutDirection.BottomToTop or TopologyLayoutDirection.RightToLeft;
            source = direction.ToString();
        }
        if (reverse) {
            report.Warn(OfficeVisioVisualDiagnosticCode.DirectionNormalized, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "layoutDirection",
                $"CFX {source} layout direction was normalized to Visio's native forward direction.");
        }
        return vertical ? VisioGraphDirection.TopToBottom : VisioGraphDirection.LeftToRight;
    }

    private static VisioGraphNodeKind MapNodeKind(VisualArtifactInterchangeNode node, bool flow) {
        if (flow) {
            return node.Flow!.Kind switch {
                FlowArtifactStepKind.Decision => VisioGraphNodeKind.Decision,
                FlowArtifactStepKind.Input or FlowArtifactStepKind.Output or FlowArtifactStepKind.Data or FlowArtifactStepKind.Document => VisioGraphNodeKind.Data,
                FlowArtifactStepKind.External or FlowArtifactStepKind.Manual => VisioGraphNodeKind.External,
                FlowArtifactStepKind.Start or FlowArtifactStepKind.End or FlowArtifactStepKind.Event => VisioGraphNodeKind.Emphasis,
                _ => VisioGraphNodeKind.Process
            };
        }
        return node.Topology!.Kind switch {
            TopologyNodeKind.Database or TopologyNodeKind.Storage or TopologyNodeKind.Queue => VisioGraphNodeKind.Data,
            TopologyNodeKind.Person or TopologyNodeKind.Team => VisioGraphNodeKind.External,
            TopologyNodeKind.Hub or TopologyNodeKind.Gateway => VisioGraphNodeKind.Emphasis,
            _ => VisioGraphNodeKind.Process
        };
    }

    private static VisioGraphConnectorKind MapEdgeKind(VisualArtifactInterchangeEdge edge, bool flow) {
        if (flow) {
            return edge.Flow!.Kind switch {
                FlowArtifactConnectorKind.Data or FlowArtifactConnectorKind.Dependency => VisioGraphConnectorKind.Data,
                FlowArtifactConnectorKind.Rejection or FlowArtifactConnectorKind.Error => VisioGraphConnectorKind.Emphasis,
                FlowArtifactConnectorKind.Retry or FlowArtifactConnectorKind.Async => VisioGraphConnectorKind.Control,
                _ => VisioGraphConnectorKind.Standard
            };
        }
        if (edge.Topology!.Status == TopologyHealthStatus.Critical) return VisioGraphConnectorKind.Emphasis;
        return edge.Topology.Kind switch {
            TopologyEdgeKind.DataFlow or TopologyEdgeKind.Dependency => VisioGraphConnectorKind.Data,
            TopologyEdgeKind.Replication or TopologyEdgeKind.Trust or TopologyEdgeKind.Mapping or
                TopologyEdgeKind.AuthenticationPath or TopologyEdgeKind.CertificateChain or
                TopologyEdgeKind.Ownership or TopologyEdgeKind.Membership => VisioGraphConnectorKind.Control,
            _ => VisioGraphConnectorKind.Standard
        };
    }

    private static VisioSequenceParticipantKind MapParticipantKind(VisualArtifactInterchangeNode node, OfficeVisioVisualConversionReport report) {
        SequenceArtifactParticipantKind kind = node.Sequence!.Kind;
        switch (kind) {
            case SequenceArtifactParticipantKind.Actor: return VisioSequenceParticipantKind.Actor;
            case SequenceArtifactParticipantKind.Boundary: return VisioSequenceParticipantKind.Boundary;
            case SequenceArtifactParticipantKind.Control: return VisioSequenceParticipantKind.Control;
            case SequenceArtifactParticipantKind.Entity: return VisioSequenceParticipantKind.Entity;
            case SequenceArtifactParticipantKind.Database: return VisioSequenceParticipantKind.Database;
            case SequenceArtifactParticipantKind.Collections:
            case SequenceArtifactParticipantKind.Queue:
                report.Warn(OfficeVisioVisualDiagnosticCode.SemanticLoss, OfficeVisioVisualEntityKind.Participant, node.Id, "participantKind",
                    $"Sequence participant kind '{kind}' was mapped to Visio's database participant shape.");
                return VisioSequenceParticipantKind.Database;
            default: return VisioSequenceParticipantKind.Participant;
        }
    }

    private static VisioSequenceMessageKind MapMessageKind(VisualArtifactInterchangeEdge edge) =>
        edge.Sequence!.Kind switch {
            SequenceArtifactMessageKind.Call => VisioSequenceMessageKind.Call,
            SequenceArtifactMessageKind.Return => VisioSequenceMessageKind.Return,
            SequenceArtifactMessageKind.Async => VisioSequenceMessageKind.Async,
            SequenceArtifactMessageKind.Event => VisioSequenceMessageKind.Event,
            _ => throw new ArgumentOutOfRangeException(nameof(edge), edge.Sequence.Kind, "Unsupported CFX sequence message kind.")
        };

    private static VisioSide MapNoteSide(SequenceArtifactNotePlacement placement, string annotationId, OfficeVisioVisualConversionReport report) {
        if (placement == SequenceArtifactNotePlacement.LeftOf) return VisioSide.Left;
        if (placement == SequenceArtifactNotePlacement.Over) {
            report.Warn(OfficeVisioVisualDiagnosticCode.NoteNormalized, OfficeVisioVisualEntityKind.Annotation, annotationId, "notePlacement",
                $"Sequence note '{annotationId}' placement '{placement}' was normalized to Visio's native right-side note placement.");
        }
        return VisioSide.Right;
    }

    private static string? EdgeDirection(VisualArtifactInterchangeEdge edge) => edge.Role switch {
        VisualArtifactInterchangeEdgeRole.TopologyEdge => edge.Topology!.Direction.ToString(),
        VisualArtifactInterchangeEdgeRole.FlowConnector => edge.Flow!.Direction.ToString(),
        VisualArtifactInterchangeEdgeRole.SequenceMessage => VisualLinkDirection.Forward.ToString(),
        _ => null
    };

    private static string? EdgeLineStyle(VisualArtifactInterchangeEdge edge) => edge.Role switch {
        VisualArtifactInterchangeEdgeRole.TopologyEdge => edge.Topology!.LineStyle.ToString(),
        VisualArtifactInterchangeEdgeRole.SequenceMessage => edge.Sequence!.LineStyle.ToString(),
        _ => null
    };
}
