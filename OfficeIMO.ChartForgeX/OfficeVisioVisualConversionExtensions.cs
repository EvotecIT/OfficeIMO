using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using global::ChartForgeX.Primitives;
using global::ChartForgeX.VisualArtifacts;
using global::ChartForgeX.Topology;
using global::ChartForgeX.VisualBlocks;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.ChartForgeX;

/// <summary>Projects CFX semantic artifacts into native editable OfficeIMO.Visio diagrams.</summary>
public static partial class OfficeVisioVisualConversionExtensions {
    /// <summary>Projects a typed CFX artifact into a native editable Visio document.</summary>
    public static OfficeVisioVisualConversionResult ToOfficeVisio(
        this VisualArtifact artifact,
        OfficeVisioVisualOptions? options = null,
        VisualArtifactRenderOptions? renderOptions = null) {
        if (artifact == null) throw new ArgumentNullException(nameof(artifact));
        OfficeVisioVisualConversionResult result = artifact.ToInterchangeEnvelope(renderOptions).ToOfficeVisio(options);
        if (renderOptions != null && renderOptions.Watermarks.Count > 0) {
            result.Report.Warn(OfficeVisioVisualDiagnosticCode.WatermarkNotProjected, OfficeVisioVisualEntityKind.Artifact, artifact.Id, "watermark",
                "CFX render watermarks are not projected into the native editable Visio page; keep the separately rendered SVG or PNG when watermark fidelity is required.");
        }
        return result;
    }

    /// <summary>Projects a validated CFX semantic envelope into a native editable Visio document.</summary>
    public static OfficeVisioVisualConversionResult ToOfficeVisio(
        this VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualOptions? options = null) {
        if (envelope == null) throw new ArgumentNullException(nameof(envelope));
        envelope.Validate();
        VisualArtifactInterchangeEnvelope validated = envelope;
        options ??= new OfficeVisioVisualOptions();

        VisioDocument document = VisioDocument.Create();
        document.Title = HasTitle(validated) ? CombineLabel(validated.Title, validated.Subtitle) : null;
        var report = new OfficeVisioVisualConversionReport {
            ArtifactKind = validated.Kind,
            SemanticFamily = validated.Family,
            NodeCount = validated.Nodes.Count,
            EdgeCount = validated.Edges.Count,
            AllProjectedObjectsEditable = true
        };
        ReportDisabledShapeDataFidelity(validated, options, report);

        switch (validated.Family) {
            case VisualArtifactInterchangeFamily.Topology:
                report.Projection = OfficeVisioVisualProjectionKind.Graph;
                BuildGraph(document, validated, options, report, flow: false);
                break;
            case VisualArtifactInterchangeFamily.Flow:
                report.Projection = OfficeVisioVisualProjectionKind.FlowGraph;
                BuildGraph(document, validated, options, report, flow: true);
                break;
            case VisualArtifactInterchangeFamily.Sequence:
                report.Projection = OfficeVisioVisualProjectionKind.Sequence;
                BuildSequence(document, validated, options, report);
                break;
            default:
                throw new NotSupportedException(
                    $"CFX artifact kind '{validated.Kind}' does not carry a supported semantic family. " +
                    "Keep the separately rendered SVG as the flat fallback or add a semantic adapter for this artifact family.");
        }

        VisioPage page = document.Pages[document.Pages.Count - 1];
        return new OfficeVisioVisualConversionResult(validated, document, page, report);
    }

    /// <summary>Parses CFX interchange JSON bytes and projects them into a native editable Visio document.</summary>
    public static OfficeVisioVisualConversionResult ToOfficeVisio(
        this byte[] interchangeJson,
        OfficeVisioVisualOptions? options = null) {
        if (interchangeJson == null) throw new ArgumentNullException(nameof(interchangeJson));
        return VisualArtifactInterchangeEnvelope.FromUtf8Json(interchangeJson).ToOfficeVisio(options);
    }

    private static void BuildGraph(
        VisioDocument document,
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report,
        bool flow) {
        List<VisioGraphNodeRecord> nodes = envelope.Nodes.Select(node => MapGraphNode(node, options, report, flow)).ToList();
        List<VisioGraphEdgeRecord> edges = envelope.Edges.Select(edge => MapGraphEdge(edge, options, report, flow)).ToList();
        List<VisioGraphClusterRecord> groups = options.IncludeGroups
            ? MapGraphGroups(envelope, options, report)
            : new List<VisioGraphClusterRecord>();
        report.GroupCount = groups.Count;
        report.AnnotationCount = 0;
        foreach (VisualArtifactInterchangeAnnotation annotation in envelope.Annotations) {
            report.Warn(OfficeVisioVisualDiagnosticCode.AnnotationNotProjected, OfficeVisioVisualEntityKind.Annotation, annotation.Id, annotation.Kind,
                $"Annotation '{annotation.Id}' of kind '{annotation.Kind}' remains in the CFX envelope but has no native graph mapping.");
        }
        if (!options.IncludeGroups && envelope.Groups.Count > 0) {
            report.Warn(OfficeVisioVisualDiagnosticCode.GroupNotProjected, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "groups",
                "Graph groups remain in the CFX envelope because native group projection was disabled by the conversion options.");
        }
        if (envelope.Extensions.Count > 0) {
            report.Warn(OfficeVisioVisualDiagnosticCode.ExtensionsNotProjected, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "extensions",
                "Artifact-level extensions remain in the CFX envelope and are not duplicated into the native Visio graph page or document.");
        }
        ReportArtifactAccessibilityFidelity(envelope, report);
        ReportArtifactPresentationFidelity(envelope, report);
        ReportScenarioFidelity(envelope, report);
        ReportGraphSemanticFidelity(envelope, report, flow);

        document.GraphDiagram(options.PageName, builder => {
            ConfigureGraph(builder, envelope, options, report, flow);
            builder.Import(nodes, edges, groups);
        });
        if (options.UseNaturalPageSize) {
            document.Pages[document.Pages.Count - 1].CenterContent();
        }
    }

    private static void ConfigureGraph(
        VisioGraphDiagramBuilder builder,
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report,
        bool flow) {
        (double width, double height) = ResolvePageSize(envelope, options);
        builder.PageSize(width, height).FitPageToGraph();
        builder.Layout(MapLayout(envelope, flow, report));
        builder.Direction(MapDirection(envelope, flow, report));
        if (envelope.Nodes.Any(node => node.X.HasValue || node.Y.HasValue || node.Width.HasValue || node.Height.HasValue) ||
            envelope.Groups.Any(group => group.X.HasValue || group.Y.HasValue || group.Width.HasValue || group.Height.HasValue)) {
            report.Warn(OfficeVisioVisualDiagnosticCode.LayoutRecomputed, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "coordinates",
                "Native Visio layout was recomputed; prepared CFX pixel coordinates and dimensions remain available in the semantic envelope.");
        }
        if (options.IncludeTitle && HasTitle(envelope)) {
            builder.Title(CombineLabel(envelope.Title, envelope.Subtitle), UniqueTitleId(envelope));
        }
    }

    private static VisioGraphNodeRecord MapGraphNode(
        VisualArtifactInterchangeNode node,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report,
        bool flow) {
        var record = new VisioGraphNodeRecord(node.Id, CombineLabel(node.Label, node.Subtitle)) {
            Kind = MapNodeKind(node, flow),
            HyperlinkAddress = options.IncludeHyperlinks ? node.Href : null,
            HyperlinkDescription = options.IncludeHyperlinks ? node.Tooltip : null,
            LineColor = MapNativeColor(node.Color, "Node", node.Id, report),
            FillColor = MapNativeColor(node.BackgroundColor, "Node background", node.Id, report)
        };
        if (options.IncludeShapeData) {
            AddCommonShapeData(record.ShapeData, node.Kind, node.Status, node.GroupId, node.Extensions, report, "node '" + node.Id + "'");
            AddValue(record.ShapeData, "CFX.Role", node.Role.ToString());
            AddValue(record.ShapeData, flow ? "CFX.FlowStepKind" : "CFX.TopologyNodeKind", flow ? node.Flow!.Kind.ToString() : node.Topology!.Kind.ToString());
            AddMetricData(record.ShapeData, node.Metrics, report, "node '" + node.Id + "'");
            AddValue(record.ShapeData, "CFX.Icon", node.IconId);
            AddValue(record.ShapeData, "CFX.Symbol", node.Symbol);
            AddValue(record.ShapeData, "CFX.Badge", node.Badge);
            AddValue(record.ShapeData, "CFX.Color", node.Color);
            AddValue(record.ShapeData, "CFX.BackgroundColor", node.BackgroundColor);
            AddDetailData(record.ShapeData, node.Details, report, "node '" + node.Id + "'");
            AddPortData(record.ShapeData, node.Ports, report, "node '" + node.Id + "'");
        }
        PreserveHyperlinkFidelity(record.ShapeData, node.Href, options, report, "Node", node.Id);
        PreserveTooltipFidelity(record.ShapeData, node.Tooltip, node.Href, options, report, "Node", node.Id);
        return record;
    }

    private static VisioGraphEdgeRecord MapGraphEdge(
        VisualArtifactInterchangeEdge edge,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report,
        bool flow) {
        var record = new VisioGraphEdgeRecord(edge.Id, edge.SourceId, edge.TargetId) {
            Kind = MapEdgeKind(edge, flow),
            Label = CombineEdgeLabel(edge),
            HyperlinkAddress = options.IncludeHyperlinks ? edge.Href : null,
            HyperlinkDescription = options.IncludeHyperlinks ? edge.Tooltip : null,
            LineStyle = MapGraphLineStyle(edge.Topology?.LineStyle),
            LineColor = MapNativeColor(edge.Color, "Edge", edge.Id, report)
        };
        ApplyGraphEdgeDirection(record, flow ? edge.Flow!.Direction : edge.Topology!.Direction);
        bool hasExplicitPortAttachment = !string.IsNullOrWhiteSpace(edge.SourcePortId) ||
                                         !string.IsNullOrWhiteSpace(edge.TargetPortId) ||
                                         (!flow && (edge.Topology!.SourcePort != global::ChartForgeX.Topology.TopologyEdgePort.Auto ||
                                                    edge.Topology.TargetPort != global::ChartForgeX.Topology.TopologyEdgePort.Auto));
        if (hasExplicitPortAttachment) {
            report.Warn(OfficeVisioVisualDiagnosticCode.PortAttachmentNormalized, OfficeVisioVisualEntityKind.Edge, edge.Id, "ports",
                $"Edge '{edge.Id}' requested CFX port attachment; native Visio graph layout selected connector sides while the original port semantics remain in the CFX envelope and, when enabled, Shape Data.");
        }
        if (!string.IsNullOrWhiteSpace(edge.SourceLabel) || !string.IsNullOrWhiteSpace(edge.TargetLabel)) {
            report.Warn(OfficeVisioVisualDiagnosticCode.EndpointLabelsNotRendered, OfficeVisioVisualEntityKind.Edge, edge.Id, "endpointLabels",
                options.IncludeShapeData
                    ? $"Edge '{edge.Id}' endpoint labels are retained as Shape Data because native Visio graph connectors do not render endpoint labels."
                    : $"Edge '{edge.Id}' endpoint labels remain only in the CFX envelope because Shape Data projection was disabled and native Visio graph connectors do not render endpoint labels.");
        }
        if (options.IncludeShapeData) {
            AddCommonShapeData(record.ShapeData, edge.Kind, edge.Status, null, edge.Extensions, report, "edge '" + edge.Id + "'");
            AddValue(record.ShapeData, "CFX.Role", edge.Role.ToString());
            AddValue(record.ShapeData, flow ? "CFX.FlowConnectorKind" : "CFX.TopologyEdgeKind", flow ? edge.Flow!.Kind.ToString() : edge.Topology!.Kind.ToString());
            AddMetricData(record.ShapeData, edge.Metrics, report, "edge '" + edge.Id + "'");
            AddValue(record.ShapeData, "CFX.Direction", EdgeDirection(edge));
            AddValue(record.ShapeData, "CFX.LineStyle", EdgeLineStyle(edge));
            if (!flow) {
                AddValue(record.ShapeData, "CFX.SourcePort", edge.Topology!.SourcePort.ToString());
                AddValue(record.ShapeData, "CFX.TargetPort", edge.Topology.TargetPort.ToString());
            }
            AddValue(record.ShapeData, "CFX.SourcePortId", edge.SourcePortId);
            AddValue(record.ShapeData, "CFX.TargetPortId", edge.TargetPortId);
            AddValue(record.ShapeData, "CFX.SourceLabel", edge.SourceLabel);
            AddValue(record.ShapeData, "CFX.TargetLabel", edge.TargetLabel);
            AddValue(record.ShapeData, "CFX.Order", edge.Order.ToString(CultureInfo.InvariantCulture));
            AddValue(record.ShapeData, "CFX.Color", edge.Color);
        }
        PreserveHyperlinkFidelity(record.ShapeData, edge.Href, options, report, "Edge", edge.Id);
        PreserveTooltipFidelity(record.ShapeData, edge.Tooltip, edge.Href, options, report, "Edge", edge.Id);
        return record;
    }

    private static List<VisioGraphClusterRecord> MapGraphGroups(
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        var groups = new List<VisioGraphClusterRecord>();
        foreach (VisualArtifactInterchangeGroup group in envelope.Groups) {
            string[] nodeIds = envelope.Nodes.Where(node => string.Equals(node.GroupId, group.Id, StringComparison.Ordinal)).Select(node => node.Id).ToArray();
            if (nodeIds.Length == 0) {
                report.Warn(OfficeVisioVisualDiagnosticCode.GroupNotProjected, OfficeVisioVisualEntityKind.Group, group.Id, "emptyGroup",
                    $"Group '{group.Id}' was not emitted because native Visio containers require at least one node.");
                continue;
            }
            var record = new VisioGraphClusterRecord(group.Id, CombineLabel(group.Label, group.Subtitle), nodeIds) {
                HyperlinkAddress = options.IncludeHyperlinks ? group.Href : null,
                HyperlinkDescription = options.IncludeHyperlinks ? group.Tooltip : null,
                LineColor = MapNativeColor(group.Color, "Group", group.Id, report)
            };
            if (options.IncludeShapeData) {
                AddCommonShapeData(record.ShapeData, group.Kind, group.Status, null, group.Extensions, report, "group '" + group.Id + "'");
                AddValue(record.ShapeData, "CFX.Role", group.Role.ToString());
                AddValue(record.ShapeData, "CFX.Color", group.Color);
            }
            PreserveHyperlinkFidelity(record.ShapeData, group.Href, options, report, "Group", group.Id);
            PreserveTooltipFidelity(record.ShapeData, group.Tooltip, group.Href, options, report, "Group", group.Id);
            groups.Add(record);
        }
        return groups;
    }

    private static void BuildSequence(
        VisioDocument document,
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        List<VisualArtifactInterchangeNode> participants = envelope.Nodes
            .OrderBy(node => node.Sequence!.Order)
            .ToList();
        if (participants.Any(participant => participant.X.HasValue || participant.Y.HasValue || participant.Width.HasValue || participant.Height.HasValue)) {
            report.Warn(OfficeVisioVisualDiagnosticCode.LayoutRecomputed, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "participantBounds",
                "Native Visio sequence layout was recomputed; prepared CFX participant coordinates and dimensions remain available in the semantic envelope.");
        }
        List<VisualArtifactInterchangeEdge> messages = envelope.Edges.OrderBy(edge => edge.Order).ToList();
        (double width, double height) = ResolveSequenceLayoutPageSize(envelope, options);
        bool includeTitle = options.IncludeTitle && HasTitle(envelope);
        var ids = new SequenceVisioIdMap(participants, messages, envelope.Annotations, includeTitle);
        IReadOnlyList<SequenceActivationProjection> activationProjections = Array.Empty<SequenceActivationProjection>();

        ReportSequenceIdMappings(participants, messages, envelope.Annotations, ids, report);

        document.SequenceDiagram(options.PageName, builder => {
            builder.PageSize(width, height);
            if (ids.TitleId != null) {
                builder.Title(CombineLabel(envelope.Title, envelope.Subtitle), ids.TitleId);
            }
            foreach (VisualArtifactInterchangeNode participant in participants) {
                builder.Participant(ids.Participant(participant.Id), CombineLabel(participant.Label, participant.Subtitle), MapParticipantKind(participant, report));
            }
            foreach (VisualArtifactInterchangeEdge message in messages) {
                VisioSequenceMessageKind kind = MapMessageKind(message);
                if (string.Equals(message.SourceId, message.TargetId, StringComparison.Ordinal)) {
                    builder.SelfMessage(ids.Participant(message.SourceId), CombineEdgeLabel(message) ?? string.Empty, kind, ids.Message(message.Id));
                } else {
                    builder.Message(ids.Participant(message.SourceId), ids.Participant(message.TargetId), CombineEdgeLabel(message) ?? string.Empty, kind, ids.Message(message.Id));
                }
            }
            activationProjections = AddSequenceActivations(builder, envelope, messages, ids, report);
            report.AnnotationCount = AddSequenceAnnotations(builder, envelope, ids, report);
        });

        VisioPage page = document.Pages[document.Pages.Count - 1];
        ApplySequenceParticipantData(page, participants, ids, options, report);
        ApplySequenceMessageData(page, messages, ids, options, report);
        ApplySequenceAnnotationData(page, envelope.Annotations, activationProjections, ids, options, report);
        if (options.UseNaturalPageSize) {
            page.Width = Math.Max(page.Width, width);
            page.Height = Math.Max(page.Height, height);
            page.CenterContent();
        } else {
            page.FitToContent(0.5D);
        }

        if (envelope.Extensions.Count > 0) {
            report.Warn(OfficeVisioVisualDiagnosticCode.ExtensionsNotProjected, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "extensions",
                options.IncludeShapeData
                    ? "Sequence-level extensions remain available in the CFX envelope; participant extensions are projected into native Visio Shape Data."
                    : "Sequence-level extensions remain only in the CFX envelope because Shape Data projection was disabled.");
        }
        if (envelope.Groups.Count > 0) {
            report.Warn(OfficeVisioVisualDiagnosticCode.GroupNotProjected, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "groups",
                "Sequence groups remain in the CFX envelope because native Visio sequence diagrams do not project graph containers.");
        }
        if (participants.Any(participant => participant.Ports.Count > 0)) {
            report.Warn(OfficeVisioVisualDiagnosticCode.PortAttachmentNormalized, OfficeVisioVisualEntityKind.Artifact, envelope.Id, "participantPorts",
                options.IncludeShapeData
                    ? "Sequence participant ports remain in CFX Shape Data and the semantic envelope because native messages attach to participant lifelines."
                    : "Sequence participant ports remain only in the CFX envelope because Shape Data projection was disabled and native messages attach to participant lifelines.");
        }
        ReportArtifactAccessibilityFidelity(envelope, report);
        ReportArtifactPresentationFidelity(envelope, report);
        ReportScenarioFidelity(envelope, report);
        ReportSequenceSemanticFidelity(envelope, report);
    }

    private static IReadOnlyList<SequenceActivationProjection> AddSequenceActivations(
        VisioSequenceDiagramBuilder builder,
        VisualArtifactInterchangeEnvelope envelope,
        IReadOnlyList<VisualArtifactInterchangeEdge> messages,
        SequenceVisioIdMap ids,
        OfficeVisioVisualConversionReport report) {
        var open = new Dictionary<string, Stack<SequenceActivationOpen>>(StringComparer.Ordinal);
        var activations = new List<SequenceActivationProjection>();
        var changes = new List<SequenceActivationChange>();
        int ordinal = 0;
        for (int index = 0; index < messages.Count; index++) {
            VisualArtifactInterchangeEdge message = messages[index];
            if (message.Sequence!.ActivatesTarget) {
                string participant = ids.Participant(message.TargetId);
                changes.Add(new SequenceActivationChange(index, 0, ordinal++, participant, active: true, message.Id));
            }
            if (message.Sequence.Deactivates) {
                string participant = ids.Participant(message.SourceId);
                changes.Add(new SequenceActivationChange(index, 0, ordinal++, participant, active: false, message.Id));
            }
        }
        foreach (VisualArtifactInterchangeAnnotation activationEvent in envelope.Annotations
                     .Where(annotation => annotation.Role == VisualArtifactInterchangeAnnotationRole.SequenceActivation)) {
            int row = activationEvent.StartIndex!.Value;
            string participant = ids.Participant(activationEvent.TargetIds[0]);
            bool active = activationEvent.Sequence!.ActivationState!.Value;
            SequenceActivationChange? duplicate = changes.FirstOrDefault(change =>
                change.Row == row && string.Equals(change.Participant, participant, StringComparison.Ordinal) && change.Active == active);
            if (duplicate != null) {
                duplicate.Annotations.Add(activationEvent);
            } else {
                var change = new SequenceActivationChange(row, 1, ordinal++, participant, active, activationEvent.Id);
                change.Annotations.Add(activationEvent);
                changes.Add(change);
            }
        }
        foreach (SequenceActivationChange change in changes.OrderBy(item => item.Row).ThenBy(item => item.SourceOrder).ThenBy(item => item.Ordinal)) {
            if (change.Active) {
                if (!open.TryGetValue(change.Participant, out Stack<SequenceActivationOpen>? starts)) {
                    starts = new Stack<SequenceActivationOpen>();
                    open.Add(change.Participant, starts);
                }
                starts.Push(new SequenceActivationOpen(change.Row, change.Annotations));
            } else if (open.TryGetValue(change.Participant, out Stack<SequenceActivationOpen>? starts) && starts.Count > 0) {
                SequenceActivationOpen start = starts.Pop();
                activations.Add(CreateActivationProjection(ids, change.Participant, start.Row, change.Row, start.Annotations, change.Annotations));
            } else {
                report.Warn(OfficeVisioVisualDiagnosticCode.SemanticLoss, OfficeVisioVisualEntityKind.Annotation, change.EntityId, "activation",
                    $"Sequence deactivation '{change.EntityId}' had no matching open activation and was not projected.");
            }
        }
        int finalRow = Math.Max(Math.Max(0, messages.Count - 1), changes.Count == 0 ? 0 : changes.Max(item => item.Row));
        foreach (KeyValuePair<string, Stack<SequenceActivationOpen>> item in open.OrderBy(pair => pair.Key, StringComparer.Ordinal)) {
            while (item.Value.Count > 0) {
                SequenceActivationOpen start = item.Value.Pop();
                activations.Add(CreateActivationProjection(ids, item.Key, start.Row, finalRow, start.Annotations, Array.Empty<VisualArtifactInterchangeAnnotation>()));
            }
        }
        foreach (SequenceActivationProjection activation in activations) {
            builder.Activation(activation.Participant, activation.Start, activation.End, activation.ShapeId);
        }
        return activations;
    }

    private static int AddSequenceAnnotations(
        VisioSequenceDiagramBuilder builder,
        VisualArtifactInterchangeEnvelope envelope,
        SequenceVisioIdMap ids,
        OfficeVisioVisualConversionReport report) {
        int projected = AddSequenceFragments(builder, envelope, ids, report);
        foreach (VisualArtifactInterchangeAnnotation annotation in envelope.Annotations) {
            int start = annotation.StartIndex ?? 0;
            if (annotation.Role == VisualArtifactInterchangeAnnotationRole.SequenceNote) {
                if (annotation.TargetIds.Count == 0) {
                    report.Warn(OfficeVisioVisualDiagnosticCode.AnnotationNotProjected, OfficeVisioVisualEntityKind.Annotation, annotation.Id, "noteTarget",
                        $"Sequence note '{annotation.Id}' has no remaining participant target and was not projected into native Visio.");
                    continue;
                }
                builder.Note(ids.Participant(annotation.TargetIds[0]), annotation.Text, start, MapNoteSide(annotation.Sequence!.NotePlacement!.Value, annotation.Id, report), ids.Annotation(annotation.Id));
                if (annotation.TargetIds.Count > 1 || annotation.Sequence.NotePlacement == SequenceArtifactNotePlacement.Over) {
                    report.Warn(OfficeVisioVisualDiagnosticCode.NoteNormalized, OfficeVisioVisualEntityKind.Annotation, annotation.Id, "notePlacement",
                        $"Sequence note '{annotation.Id}' was attached to its first participant because native side notes do not span multiple participants.");
                }
                projected++;
                continue;
            }
            if (annotation.Role is VisualArtifactInterchangeAnnotationRole.SequenceActivation or
                VisualArtifactInterchangeAnnotationRole.SequenceBlock or
                VisualArtifactInterchangeAnnotationRole.SequenceBranch) continue;
            report.Warn(OfficeVisioVisualDiagnosticCode.AnnotationNotProjected, OfficeVisioVisualEntityKind.Annotation, annotation.Id, annotation.Role.ToString(),
                $"Sequence annotation '{annotation.Id}' of role '{annotation.Role}' remains in the CFX envelope but has no native Visio mapping.");
        }
        return projected;
    }

    private static (double Width, double Height) ResolvePageSize(VisualArtifactInterchangeEnvelope envelope, OfficeVisioVisualOptions options) {
        if (!options.UseNaturalPageSize) return (1D, 1D);
        double width = envelope.Width.HasValue ? envelope.Width.Value / options.PixelsPerInch : 11D;
        double height = envelope.Height.HasValue ? envelope.Height.Value / options.PixelsPerInch : 8.5D;
        return (width, height);
    }

    private static (double Width, double Height) ResolveSequenceLayoutPageSize(VisualArtifactInterchangeEnvelope envelope, OfficeVisioVisualOptions options) {
        if (!options.UseNaturalPageSize) return (1D, 1D);
        double width = envelope.Width.HasValue ? envelope.Width.Value / options.PixelsPerInch : 11D;
        double height = envelope.Height.HasValue ? envelope.Height.Value / options.PixelsPerInch : 8.5D;
        return (width, height);
    }

    private static void ApplySequenceParticipantData(
        VisioPage page,
        IEnumerable<VisualArtifactInterchangeNode> participants,
        SequenceVisioIdMap ids,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        foreach (VisualArtifactInterchangeNode participant in participants) {
            VisioShape shape = page.Shapes.Single(item => string.Equals(item.Id, ids.Participant(participant.Id), StringComparison.Ordinal));
            OfficeIMO.Drawing.OfficeColor? lineColor = MapNativeColor(participant.Color, "Sequence participant", participant.Id, report);
            OfficeIMO.Drawing.OfficeColor? fillColor = MapNativeColor(participant.BackgroundColor, "Sequence participant background", participant.Id, report);
            if (lineColor.HasValue) shape.LineColor = lineColor.Value;
            if (fillColor.HasValue) shape.FillColor = fillColor.Value;
            shape.Data["CFX.Id"] = participant.Id;
            var data = new Dictionary<string, string?>(StringComparer.Ordinal);
            if (options.IncludeShapeData) {
                AddValue(data, "CFX.Id", participant.Id);
                AddCommonShapeData(data, participant.Kind, participant.Status, participant.GroupId, participant.Extensions, report, "sequence participant '" + participant.Id + "'");
                AddValue(data, "CFX.Role", participant.Role.ToString());
                AddValue(data, "CFX.SequenceParticipantKind", participant.Sequence!.Kind.ToString());
                AddValue(data, "CFX.SequenceParticipantOrder", participant.Sequence.Order.ToString(CultureInfo.InvariantCulture));
                AddValue(data, "CFX.SequenceParticipantImplicit", participant.Sequence.IsImplicit.ToString(CultureInfo.InvariantCulture));
                AddMetricData(data, participant.Metrics, report, "sequence participant '" + participant.Id + "'");
                AddValue(data, "CFX.Icon", participant.IconId);
                AddValue(data, "CFX.Symbol", participant.Symbol);
                AddValue(data, "CFX.Badge", participant.Badge);
                AddValue(data, "CFX.Color", participant.Color);
                AddValue(data, "CFX.BackgroundColor", participant.BackgroundColor);
                AddDetailData(data, participant.Details, report, "sequence participant '" + participant.Id + "'");
                AddPortData(data, participant.Ports, report, "sequence participant '" + participant.Id + "'");
            }
            PreserveHyperlinkFidelity(data, participant.Href, options, report, "Sequence participant", participant.Id);
            PreserveTooltipFidelity(data, participant.Tooltip, participant.Href, options, report, "Sequence participant", participant.Id);
            foreach (KeyValuePair<string, string?> item in data) shape.SetShapeData(item.Key, item.Value);
            if (options.IncludeHyperlinks && !string.IsNullOrWhiteSpace(participant.Href)) {
                shape.AddHyperlink(participant.Href!, participant.Tooltip);
            }
        }
    }

    private static void ApplySequenceMessageData(
        VisioPage page,
        IEnumerable<VisualArtifactInterchangeEdge> messages,
        SequenceVisioIdMap ids,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        foreach (VisualArtifactInterchangeEdge message in messages) {
            VisioConnector connector = page.Connectors.Single(item => string.Equals(item.Id, ids.Message(message.Id), StringComparison.Ordinal));
            connector.LinePattern = OfficeStrokeDashStyleMapper.ToVisioLinePattern(MapSequenceLineStyle(message.Sequence!.LineStyle));
            OfficeIMO.Drawing.OfficeColor? lineColor = MapNativeColor(message.Color, "Sequence message", message.Id, report);
            if (lineColor.HasValue) connector.LineColor = lineColor.Value;
            connector.Data["CFX.Id"] = message.Id;
            var data = new Dictionary<string, string?>(StringComparer.Ordinal);
            if (options.IncludeShapeData) {
                AddValue(data, "CFX.Id", message.Id);
                AddCommonShapeData(data, message.Kind, message.Status, null, message.Extensions, report, "sequence message '" + message.Id + "'");
                AddValue(data, "CFX.Role", message.Role.ToString());
                AddValue(data, "CFX.SequenceMessageKind", message.Sequence!.Kind.ToString());
                AddValue(data, "CFX.SequenceActivatesTarget", message.Sequence.ActivatesTarget.ToString(CultureInfo.InvariantCulture));
                AddValue(data, "CFX.SequenceDeactivates", message.Sequence.Deactivates.ToString(CultureInfo.InvariantCulture));
                AddMetricData(data, message.Metrics, report, "sequence message '" + message.Id + "'");
                AddValue(data, "CFX.Direction", EdgeDirection(message));
                AddValue(data, "CFX.LineStyle", EdgeLineStyle(message));
                AddValue(data, "CFX.SourcePortId", message.SourcePortId);
                AddValue(data, "CFX.TargetPortId", message.TargetPortId);
                AddValue(data, "CFX.SourceLabel", message.SourceLabel);
                AddValue(data, "CFX.TargetLabel", message.TargetLabel);
                AddValue(data, "CFX.Order", message.Order.ToString(CultureInfo.InvariantCulture));
                AddValue(data, "CFX.Color", message.Color);
            }
            PreserveHyperlinkFidelity(data, message.Href, options, report, "Sequence message", message.Id);
            PreserveTooltipFidelity(data, message.Tooltip, message.Href, options, report, "Sequence message", message.Id);
            foreach (KeyValuePair<string, string?> item in data) connector.SetShapeData(item.Key, item.Value);
            if (options.IncludeHyperlinks && !string.IsNullOrWhiteSpace(message.Href)) {
                connector.AddHyperlink(message.Href!, message.Tooltip);
            }
        }
    }

    private static void ApplySequenceAnnotationData(
        VisioPage page,
        IEnumerable<VisualArtifactInterchangeAnnotation> annotations,
        IReadOnlyList<SequenceActivationProjection> activationProjections,
        SequenceVisioIdMap ids,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        foreach (VisualArtifactInterchangeAnnotation annotation in annotations.Where(item => item.Role != VisualArtifactInterchangeAnnotationRole.SequenceActivation)) {
            string nativeId = ids.Annotation(annotation.Id);
            VisioShape? shape = page.Shapes.FirstOrDefault(item =>
                string.Equals(item.Id, nativeId, StringComparison.Ordinal) ||
                string.Equals(item.GetUserCellValue("OfficeIMO.SequenceFragmentOperandId"), nativeId, StringComparison.Ordinal));
            if (shape == null) continue;
            ApplySequenceAnnotationShapeData(shape, annotation, options, report);
        }
        foreach (SequenceActivationProjection projection in activationProjections) {
            VisioShape? shape = page.Shapes.FirstOrDefault(item => string.Equals(item.Id, projection.ShapeId, StringComparison.Ordinal));
            if (shape == null) continue;
            ApplySequenceActivationShapeData(shape, projection, options, report);
        }
    }

    private static void ReportSequenceIdMappings(
        IEnumerable<VisualArtifactInterchangeNode> participants,
        IEnumerable<VisualArtifactInterchangeEdge> messages,
        IEnumerable<VisualArtifactInterchangeAnnotation> annotations,
        SequenceVisioIdMap ids,
        OfficeVisioVisualConversionReport report) {
        foreach (VisualArtifactInterchangeNode participant in participants) {
            ReportSequenceIdMapping(OfficeVisioVisualEntityKind.Participant, participant.Id, ids.Participant(participant.Id), report);
        }
        foreach (VisualArtifactInterchangeEdge message in messages) {
            ReportSequenceIdMapping(OfficeVisioVisualEntityKind.Message, message.Id, ids.Message(message.Id), report);
        }
        foreach (VisualArtifactInterchangeAnnotation annotation in annotations) {
            ReportSequenceIdMapping(OfficeVisioVisualEntityKind.Annotation, annotation.Id, ids.Annotation(annotation.Id), report);
        }
    }

    private static void ReportSequenceIdMapping(OfficeVisioVisualEntityKind entityKind, string sourceId, string visioId, OfficeVisioVisualConversionReport report) {
        if (!string.Equals(sourceId, visioId, StringComparison.Ordinal)) {
            report.Info(OfficeVisioVisualDiagnosticCode.IdRemapped, entityKind, sourceId, "id",
                $"Sequence {entityKind.ToString().ToLowerInvariant()} id '{sourceId}' was projected as '{visioId}' to avoid a collision with native Visio helper shapes.");
        }
    }

    private static void ApplyGraphEdgeDirection(VisioGraphEdgeRecord record, VisualLinkDirection direction) {
        record.Directed = direction != VisualLinkDirection.None;
        record.BeginArrow = direction is VisualLinkDirection.Backward or VisualLinkDirection.Bidirectional ? EndArrow.Triangle : EndArrow.None;
        record.EndArrow = direction is VisualLinkDirection.Forward or VisualLinkDirection.Bidirectional ? EndArrow.Triangle : EndArrow.None;
    }

    private static string UniqueTitleId(VisualArtifactInterchangeEnvelope envelope) {
        string candidate = "cfx-title";
        var ids = new HashSet<string>(envelope.Nodes.Select(node => node.Id), StringComparer.Ordinal);
        foreach (VisualArtifactInterchangeGroup group in envelope.Groups) ids.Add(group.Id);
        foreach (VisualArtifactInterchangeEdge edge in envelope.Edges) ids.Add(edge.Id);
        foreach (VisualArtifactInterchangeAnnotation annotation in envelope.Annotations) ids.Add(annotation.Id);
        while (ids.Contains(candidate)) candidate += "-title";
        return candidate;
    }

    private static string CombineLabel(string primary, string? secondary) {
        if (string.IsNullOrWhiteSpace(primary)) return secondary ?? string.Empty;
        return string.IsNullOrWhiteSpace(secondary) ? primary : primary + Environment.NewLine + secondary;
    }

    private static bool HasTitle(VisualArtifactInterchangeEnvelope envelope) =>
        !string.IsNullOrWhiteSpace(envelope.Title) || !string.IsNullOrWhiteSpace(envelope.Subtitle);

    private static string? CombineEdgeLabel(VisualArtifactInterchangeEdge edge) {
        string[] labels = new[] { edge.Label, edge.SecondaryLabel, edge.TertiaryLabel }
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value!)
            .ToArray();
        return labels.Length == 0 ? null : string.Join(" | ", labels);
    }

}
