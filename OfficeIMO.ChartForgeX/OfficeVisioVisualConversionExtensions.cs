using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using global::ChartForgeX.VisualArtifacts;
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
            result.Report.Warn("CFX render watermarks are not projected into the native editable Visio page; keep the separately rendered SVG or PNG when watermark fidelity is required.");
        }
        return result;
    }

    /// <summary>Projects a validated CFX semantic envelope into a native editable Visio document.</summary>
    public static OfficeVisioVisualConversionResult ToOfficeVisio(
        this VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualOptions? options = null) {
        if (envelope == null) throw new ArgumentNullException(nameof(envelope));
        // Round-tripping applies the same public validation used at process and ALC boundaries.
        VisualArtifactInterchangeEnvelope validated = VisualArtifactInterchangeEnvelope.FromJson(envelope.ToJson());
        options ??= new OfficeVisioVisualOptions();

        VisioDocument document = VisioDocument.Create();
        document.Title = HasTitle(validated) ? CombineLabel(validated.Title, validated.Subtitle) : null;
        var report = new OfficeVisioVisualConversionReport {
            ArtifactKind = validated.Kind,
            NodeCount = validated.Nodes.Count,
            EdgeCount = validated.Edges.Count,
            IsNativeEditable = true
        };

        switch (validated.Kind) {
            case VisualArtifactKind.Topology:
                report.Projection = nameof(VisioGraphDiagramBuilder);
                BuildGraph(document, validated, options, report, flow: false);
                break;
            case VisualArtifactKind.Flow:
                report.Projection = nameof(VisioGraphDiagramBuilder) + " (flow)";
                BuildGraph(document, validated, options, report, flow: true);
                break;
            case VisualArtifactKind.Sequence:
                report.Projection = nameof(VisioSequenceDiagramBuilder);
                BuildSequence(document, validated, options, report);
                break;
            default:
                throw new NotSupportedException(
                    $"CFX artifact kind '{validated.Kind}' does not have a native editable Visio projection. " +
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
            report.Warn($"Annotation '{annotation.Id}' of kind '{annotation.Kind}' remains in the CFX envelope but has no native graph mapping.");
        }
        if (!options.IncludeGroups && envelope.Groups.Count > 0) {
            report.Warn("Graph groups remain in the CFX envelope because native group projection was disabled by the conversion options.");
        }
        if (envelope.Metadata.Count > 0) {
            report.Warn("Artifact-level metadata remains in the CFX envelope and is not duplicated into the native Visio graph page or document.");
        }
        ReportArtifactAccessibilityFidelity(envelope, report);

        document.GraphDiagram(options.PageName, builder => {
            ConfigureGraph(builder, envelope, options, report);
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
        OfficeVisioVisualConversionReport report) {
        (double width, double height) = ResolvePageSize(envelope, options);
        builder.PageSize(width, height).FitPageToGraph();
        builder.Layout(MapLayout(envelope.Layout, report));
        builder.Direction(MapDirection(envelope.Direction, report));
        if (envelope.Nodes.Any(node => node.X.HasValue || node.Y.HasValue || node.Width.HasValue || node.Height.HasValue) ||
            envelope.Groups.Any(group => group.X.HasValue || group.Y.HasValue || group.Width.HasValue || group.Height.HasValue)) {
            report.Warn("Native Visio layout was recomputed; prepared CFX pixel coordinates and dimensions remain available in the semantic envelope.");
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
            Kind = MapNodeKind(node.Kind, flow),
            HyperlinkAddress = options.IncludeHyperlinks ? node.Href : null,
            HyperlinkDescription = options.IncludeHyperlinks ? node.Tooltip : null,
            LineColor = MapNativeColor(node.Color, "Node", node.Id, report),
            FillColor = MapNativeColor(node.BackgroundColor, "Node background", node.Id, report)
        };
        if (options.IncludeShapeData) {
            AddCommonShapeData(record.ShapeData, node.Kind, node.Status, node.GroupId, node.Metadata, report, "node '" + node.Id + "'");
            AddValue(record.ShapeData, "CFX.Icon", node.IconId);
            AddValue(record.ShapeData, "CFX.Symbol", node.Symbol);
            AddValue(record.ShapeData, "CFX.Badge", node.Badge);
            AddValue(record.ShapeData, "CFX.Color", node.Color);
            AddValue(record.ShapeData, "CFX.BackgroundColor", node.BackgroundColor);
            AddDetailData(record.ShapeData, node.Details, report, "node '" + node.Id + "'");
            AddPortData(record.ShapeData, node.Ports, report, "node '" + node.Id + "'");
        }
        PreserveTooltipFidelity(record.ShapeData, node.Tooltip, node.Href, options, report, "Node", node.Id);
        return record;
    }

    private static VisioGraphEdgeRecord MapGraphEdge(
        VisualArtifactInterchangeEdge edge,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report,
        bool flow) {
        var record = new VisioGraphEdgeRecord(edge.Id, edge.SourceId, edge.TargetId) {
            Kind = MapEdgeKind(edge.Kind, edge.Status, flow),
            Label = CombineEdgeLabel(edge),
            HyperlinkAddress = options.IncludeHyperlinks ? edge.Href : null,
            HyperlinkDescription = options.IncludeHyperlinks ? edge.Tooltip : null,
            LinePattern = MapGraphLinePattern(edge.LineStyle, edge.Id, report),
            LineColor = MapNativeColor(edge.Color, "Edge", edge.Id, report)
        };
        ApplyGraphEdgeDirection(record, edge.Direction, edge.Id, report);
        if (!string.IsNullOrWhiteSpace(edge.SourcePortId) || !string.IsNullOrWhiteSpace(edge.TargetPortId) ||
            !string.IsNullOrWhiteSpace(edge.SourcePort) || !string.IsNullOrWhiteSpace(edge.TargetPort)) {
            report.Warn($"Edge '{edge.Id}' requested CFX port attachment; native Visio graph layout selected connector sides while the original port semantics remain in the CFX envelope and, when enabled, Shape Data.");
        }
        if (options.IncludeShapeData) {
            AddCommonShapeData(record.ShapeData, edge.Kind, edge.Status, null, edge.Metadata, report, "edge '" + edge.Id + "'");
            AddValue(record.ShapeData, "CFX.Direction", edge.Direction);
            AddValue(record.ShapeData, "CFX.LineStyle", edge.LineStyle);
            AddValue(record.ShapeData, "CFX.SourcePort", edge.SourcePort);
            AddValue(record.ShapeData, "CFX.TargetPort", edge.TargetPort);
            AddValue(record.ShapeData, "CFX.SourcePortId", edge.SourcePortId);
            AddValue(record.ShapeData, "CFX.TargetPortId", edge.TargetPortId);
            AddValue(record.ShapeData, "CFX.SourceLabel", edge.SourceLabel);
            AddValue(record.ShapeData, "CFX.TargetLabel", edge.TargetLabel);
            AddValue(record.ShapeData, "CFX.Order", edge.Order.ToString(CultureInfo.InvariantCulture));
            AddValue(record.ShapeData, "CFX.Color", edge.Color);
        }
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
                report.Warn($"Group '{group.Id}' was not emitted because native Visio containers require at least one node.");
                continue;
            }
            var record = new VisioGraphClusterRecord(group.Id, CombineLabel(group.Label, group.Subtitle), nodeIds) {
                HyperlinkAddress = options.IncludeHyperlinks ? group.Href : null,
                HyperlinkDescription = options.IncludeHyperlinks ? group.Tooltip : null,
                LineColor = MapNativeColor(group.Color, "Group", group.Id, report)
            };
            if (options.IncludeShapeData) {
                AddCommonShapeData(record.ShapeData, group.Kind, group.Status, null, group.Metadata, report, "group '" + group.Id + "'");
                AddValue(record.ShapeData, "CFX.Color", group.Color);
            }
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
            .OrderBy(node => ReadInt(node.Metadata, "sequence.order", int.MaxValue))
            .ToList();
        List<VisualArtifactInterchangeEdge> messages = envelope.Edges.OrderBy(edge => edge.Order).ToList();
        (double width, double height) = ResolveSequenceLayoutPageSize(envelope, options);
        bool includeTitle = options.IncludeTitle && HasTitle(envelope);
        var ids = new SequenceVisioIdMap(participants, messages, envelope.Annotations, includeTitle);

        ReportSequenceIdMappings(participants, messages, envelope.Annotations, ids, report);

        document.SequenceDiagram(options.PageName, builder => {
            builder.PageSize(width, height);
            if (ids.TitleId != null) {
                builder.Title(CombineLabel(envelope.Title, envelope.Subtitle), ids.TitleId);
            }
            foreach (VisualArtifactInterchangeNode participant in participants) {
                builder.Participant(ids.Participant(participant.Id), CombineLabel(participant.Label, participant.Subtitle), MapParticipantKind(participant.Kind, report));
            }
            foreach (VisualArtifactInterchangeEdge message in messages) {
                VisioSequenceMessageKind kind = MapMessageKind(message);
                if (string.Equals(message.SourceId, message.TargetId, StringComparison.Ordinal)) {
                    builder.SelfMessage(ids.Participant(message.SourceId), CombineEdgeLabel(message) ?? string.Empty, kind, ids.Message(message.Id));
                } else {
                    builder.Message(ids.Participant(message.SourceId), ids.Participant(message.TargetId), CombineEdgeLabel(message) ?? string.Empty, kind, ids.Message(message.Id));
                }
            }
            AddSequenceActivations(builder, messages, ids);
            report.AnnotationCount = AddSequenceAnnotations(builder, envelope, ids, report);
        });

        VisioPage page = document.Pages[document.Pages.Count - 1];
        ApplySequenceParticipantData(page, participants, ids, options, report);
        ApplySequenceMessageData(page, messages, ids, options, report);
        ApplySequenceAnnotationData(page, envelope.Annotations, ids, options, report);
        if (options.UseNaturalPageSize) {
            page.Width = Math.Max(page.Width, width);
            page.Height = Math.Max(page.Height, height);
            page.CenterContent();
        } else {
            page.FitToContent(0.5D);
        }

        if (envelope.Metadata.Count > 0) {
            report.Warn(options.IncludeShapeData
                ? "Sequence-level metadata remains available in the CFX envelope; participant metadata is projected into native Visio Shape Data."
                : "Sequence-level metadata remains only in the CFX envelope because Shape Data projection was disabled.");
        }
        if (envelope.Groups.Count > 0) {
            report.Warn("Sequence groups remain in the CFX envelope because native Visio sequence diagrams do not project graph containers.");
        }
        if (participants.Any(participant => participant.Ports.Count > 0)) {
            report.Warn("Sequence participant ports remain in CFX Shape Data and the semantic envelope because native messages attach to participant lifelines.");
        }
        ReportArtifactAccessibilityFidelity(envelope, report);
    }

    private static void AddSequenceActivations(
        VisioSequenceDiagramBuilder builder,
        IReadOnlyList<VisualArtifactInterchangeEdge> messages,
        SequenceVisioIdMap ids) {
        var open = new Dictionary<string, Stack<int>>(StringComparer.Ordinal);
        var activations = new List<(string Participant, int Start, int End)>();
        for (int index = 0; index < messages.Count; index++) {
            VisualArtifactInterchangeEdge message = messages[index];
            if (ReadBool(message.Metadata, "sequence.activatesTarget")) {
                string targetId = ids.Participant(message.TargetId);
                if (!open.TryGetValue(targetId, out Stack<int>? starts)) {
                    starts = new Stack<int>();
                    open.Add(targetId, starts);
                }
                starts.Push(index);
            }
            string sourceId = ids.Participant(message.SourceId);
            if (ReadBool(message.Metadata, "sequence.deactivates") && open.TryGetValue(sourceId, out Stack<int>? sourceStarts) && sourceStarts.Count > 0) {
                activations.Add((sourceId, sourceStarts.Pop(), index));
            }
        }
        int finalRow = Math.Max(0, messages.Count - 1);
        foreach (KeyValuePair<string, Stack<int>> item in open.OrderBy(pair => pair.Key, StringComparer.Ordinal)) {
            while (item.Value.Count > 0) activations.Add((item.Key, item.Value.Pop(), finalRow));
        }
        for (int index = 0; index < activations.Count; index++) {
            var activation = activations[index];
            builder.Activation(activation.Participant, activation.Start, activation.End, ids.Activation());
        }
    }

    private static int AddSequenceAnnotations(
        VisioSequenceDiagramBuilder builder,
        VisualArtifactInterchangeEnvelope envelope,
        SequenceVisioIdMap ids,
        OfficeVisioVisualConversionReport report) {
        int projected = 0;
        foreach (VisualArtifactInterchangeAnnotation annotation in envelope.Annotations) {
            int start = annotation.StartIndex ?? 0;
            int end = annotation.EndIndex ?? start;
            if (annotation.Kind.StartsWith("SequenceBlock:", StringComparison.Ordinal)) {
                string blockKind = annotation.Kind.Substring("SequenceBlock:".Length);
                builder.Fragment(CombineLabel(blockKind, annotation.Text), start, end, annotation.TargetIds.Select(ids.Participant), ids.Annotation(annotation.Id));
                projected++;
                continue;
            }
            if (string.Equals(annotation.Kind, "SequenceNote", StringComparison.Ordinal) && annotation.TargetIds.Count > 0) {
                builder.Note(ids.Participant(annotation.TargetIds[0]), annotation.Text, start, MapNoteSide(annotation.Placement), ids.Annotation(annotation.Id));
                if (annotation.TargetIds.Count > 1 || string.Equals(annotation.Placement, "Over", StringComparison.OrdinalIgnoreCase)) {
                    report.Warn($"Sequence note '{annotation.Id}' was attached to its first participant because native side notes do not span multiple participants.");
                }
                projected++;
                continue;
            }
            report.Warn($"Sequence annotation '{annotation.Id}' of kind '{annotation.Kind}' remains in the CFX envelope but has no native Visio mapping.");
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
                AddCommonShapeData(data, participant.Kind, participant.Status, participant.GroupId, participant.Metadata, report, "sequence participant '" + participant.Id + "'");
                AddValue(data, "CFX.Icon", participant.IconId);
                AddValue(data, "CFX.Symbol", participant.Symbol);
                AddValue(data, "CFX.Badge", participant.Badge);
                AddValue(data, "CFX.Color", participant.Color);
                AddValue(data, "CFX.BackgroundColor", participant.BackgroundColor);
                AddDetailData(data, participant.Details, report, "sequence participant '" + participant.Id + "'");
                AddPortData(data, participant.Ports, report, "sequence participant '" + participant.Id + "'");
            }
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
            ApplySequenceMessageDirection(connector, message.Direction, message.Id, report);
            int? linePattern = MapSequenceLinePattern(message.LineStyle, message.Id, report);
            if (linePattern.HasValue) connector.LinePattern = linePattern.Value;
            OfficeIMO.Drawing.OfficeColor? lineColor = MapNativeColor(message.Color, "Sequence message", message.Id, report);
            if (lineColor.HasValue) connector.LineColor = lineColor.Value;
            connector.Data["CFX.Id"] = message.Id;
            var data = new Dictionary<string, string?>(StringComparer.Ordinal);
            if (options.IncludeShapeData) {
                AddValue(data, "CFX.Id", message.Id);
                AddCommonShapeData(data, message.Kind, message.Status, null, message.Metadata, report, "sequence message '" + message.Id + "'");
                AddValue(data, "CFX.Direction", message.Direction);
                AddValue(data, "CFX.LineStyle", message.LineStyle);
                AddValue(data, "CFX.SourcePort", message.SourcePort);
                AddValue(data, "CFX.TargetPort", message.TargetPort);
                AddValue(data, "CFX.SourcePortId", message.SourcePortId);
                AddValue(data, "CFX.TargetPortId", message.TargetPortId);
                AddValue(data, "CFX.SourceLabel", message.SourceLabel);
                AddValue(data, "CFX.TargetLabel", message.TargetLabel);
                AddValue(data, "CFX.Order", message.Order.ToString(CultureInfo.InvariantCulture));
                AddValue(data, "CFX.Color", message.Color);
            }
            PreserveTooltipFidelity(data, message.Tooltip, message.Href, options, report, "Sequence message", message.Id);
            foreach (KeyValuePair<string, string?> item in data) connector.SetShapeData(item.Key, item.Value);
            if (options.IncludeHyperlinks && !string.IsNullOrWhiteSpace(message.Href)) {
                connector.AddHyperlink(message.Href!, message.Tooltip);
            }
        }
    }

    private static void ApplySequenceMessageDirection(
        VisioConnector connector,
        string? direction,
        string messageId,
        OfficeVisioVisualConversionReport report) {
        EndArrow arrow = connector.EndArrow ?? EndArrow.Triangle;
        if (arrow == EndArrow.None) arrow = EndArrow.Triangle;
        connector.BeginArrow = EndArrow.None;
        connector.EndArrow = arrow;
        if (string.IsNullOrWhiteSpace(direction) || string.Equals(direction, "Forward", StringComparison.OrdinalIgnoreCase)) return;
        if (string.Equals(direction, "None", StringComparison.OrdinalIgnoreCase)) {
            connector.EndArrow = EndArrow.None;
            return;
        }
        if (string.Equals(direction, "Backward", StringComparison.OrdinalIgnoreCase)) {
            connector.BeginArrow = arrow;
            connector.EndArrow = EndArrow.None;
            return;
        }
        if (string.Equals(direction, "Bidirectional", StringComparison.OrdinalIgnoreCase)) {
            connector.BeginArrow = arrow;
            return;
        }
        report.Warn($"Sequence message '{messageId}' direction '{direction}' was normalized to a forward Visio connector.");
    }

    private static void ApplySequenceAnnotationData(
        VisioPage page,
        IEnumerable<VisualArtifactInterchangeAnnotation> annotations,
        SequenceVisioIdMap ids,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        foreach (VisualArtifactInterchangeAnnotation annotation in annotations) {
            VisioShape? shape = page.Shapes.FirstOrDefault(item => string.Equals(item.Id, ids.Annotation(annotation.Id), StringComparison.Ordinal));
            if (shape == null) continue;
            shape.Data["CFX.Id"] = annotation.Id;
            if (!options.IncludeShapeData) continue;
            var data = new Dictionary<string, string?>(StringComparer.Ordinal);
            AddValue(data, "CFX.Id", annotation.Id);
            AddCommonShapeData(data, annotation.Kind, null, null, annotation.Metadata, report, "sequence annotation '" + annotation.Id + "'");
            AddValue(data, "CFX.Placement", annotation.Placement);
            AddValue(data, "CFX.TargetIds", annotation.TargetIds.Count == 0 ? null : string.Join(",", annotation.TargetIds));
            AddValue(data, "CFX.StartIndex", annotation.StartIndex?.ToString(CultureInfo.InvariantCulture));
            AddValue(data, "CFX.EndIndex", annotation.EndIndex?.ToString(CultureInfo.InvariantCulture));
            foreach (KeyValuePair<string, string?> item in data) shape.SetShapeData(item.Key, item.Value);
        }
    }

    private static void ReportSequenceIdMappings(
        IEnumerable<VisualArtifactInterchangeNode> participants,
        IEnumerable<VisualArtifactInterchangeEdge> messages,
        IEnumerable<VisualArtifactInterchangeAnnotation> annotations,
        SequenceVisioIdMap ids,
        OfficeVisioVisualConversionReport report) {
        foreach (VisualArtifactInterchangeNode participant in participants) {
            ReportSequenceIdMapping("participant", participant.Id, ids.Participant(participant.Id), report);
        }
        foreach (VisualArtifactInterchangeEdge message in messages) {
            ReportSequenceIdMapping("message", message.Id, ids.Message(message.Id), report);
        }
        foreach (VisualArtifactInterchangeAnnotation annotation in annotations) {
            ReportSequenceIdMapping("annotation", annotation.Id, ids.Annotation(annotation.Id), report);
        }
    }

    private static void ReportSequenceIdMapping(string kind, string sourceId, string visioId, OfficeVisioVisualConversionReport report) {
        if (!string.Equals(sourceId, visioId, StringComparison.Ordinal)) {
            report.Warn($"Sequence {kind} id '{sourceId}' was projected as '{visioId}' to avoid a collision with native Visio helper shapes.");
        }
    }

    private static void ApplyGraphEdgeDirection(
        VisioGraphEdgeRecord record,
        string? direction,
        string edgeId,
        OfficeVisioVisualConversionReport report) {
        record.Directed = true;
        record.BeginArrow = EndArrow.None;
        record.EndArrow = EndArrow.Triangle;
        if (string.IsNullOrWhiteSpace(direction) || string.Equals(direction, "Forward", StringComparison.OrdinalIgnoreCase)) return;
        if (string.Equals(direction, "None", StringComparison.OrdinalIgnoreCase)) {
            record.Directed = false;
            record.EndArrow = EndArrow.None;
            return;
        }
        if (string.Equals(direction, "Backward", StringComparison.OrdinalIgnoreCase)) {
            record.BeginArrow = EndArrow.Triangle;
            record.EndArrow = EndArrow.None;
            return;
        }
        if (string.Equals(direction, "Bidirectional", StringComparison.OrdinalIgnoreCase)) {
            record.BeginArrow = EndArrow.Triangle;
            return;
        }
        report.Warn($"Edge '{edgeId}' direction '{direction}' was normalized to a forward Visio connector.");
    }

    private static VisioGraphLayout MapLayout(string layout, OfficeVisioVisualConversionReport report) {
        if (string.IsNullOrWhiteSpace(layout) || string.Equals(layout, "Layered", StringComparison.OrdinalIgnoreCase)) return VisioGraphLayout.Layered;
        if (string.Equals(layout, "Dense", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(layout, "Grid", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(layout, "GroupGrid", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(layout, "DenseGrouped", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(layout, "Matrix", StringComparison.OrdinalIgnoreCase)) return VisioGraphLayout.Grid;
        if (string.Equals(layout, "Force", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(layout, "Radial", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(layout, "ForceDirected", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(layout, "RelationshipRadial", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(layout, "HubAndSpoke", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(layout, "MindMap", StringComparison.OrdinalIgnoreCase)) return VisioGraphLayout.Radial;
        if (!string.Equals(layout, "Manual", StringComparison.OrdinalIgnoreCase)) {
            report.Warn($"CFX graph layout '{layout}' was normalized to Visio's native layered layout.");
        }
        return VisioGraphLayout.Layered;
    }

    private static VisioGraphDirection MapDirection(string direction, OfficeVisioVisualConversionReport report) {
        if (string.Equals(direction, "TopToBottom", StringComparison.OrdinalIgnoreCase) || string.Equals(direction, "BottomToTop", StringComparison.OrdinalIgnoreCase)) {
            if (string.Equals(direction, "BottomToTop", StringComparison.OrdinalIgnoreCase)) report.Warn("Bottom-to-top direction was normalized to Visio's native top-to-bottom graph layout.");
            return VisioGraphDirection.TopToBottom;
        }
        if (string.Equals(direction, "RightToLeft", StringComparison.OrdinalIgnoreCase)) report.Warn("Right-to-left direction was normalized to Visio's native left-to-right graph layout.");
        return VisioGraphDirection.LeftToRight;
    }

    private static VisioGraphNodeKind MapNodeKind(string kind, bool flow) {
        if (Contains(kind, "Decision")) return VisioGraphNodeKind.Decision;
        if (Contains(kind, "Data") || Contains(kind, "Database") || Contains(kind, "Store")) return VisioGraphNodeKind.Data;
        if (Contains(kind, "External") || Contains(kind, "Actor")) return VisioGraphNodeKind.External;
        if (Contains(kind, "Critical") || Contains(kind, "Emphasis") || (flow && (Contains(kind, "Start") || Contains(kind, "End")))) return VisioGraphNodeKind.Emphasis;
        return VisioGraphNodeKind.Process;
    }

    private static VisioGraphConnectorKind MapEdgeKind(string kind, string? status, bool flow) {
        if (Contains(kind, "Data") || Contains(kind, "Dependency")) return VisioGraphConnectorKind.Data;
        if (Contains(kind, "Control") || Contains(kind, "Retry") || Contains(kind, "Async")) return VisioGraphConnectorKind.Control;
        if (Contains(kind, "Error") || Contains(kind, "Reject") || Contains(status, "Critical") || Contains(status, "Error")) return VisioGraphConnectorKind.Emphasis;
        return flow ? VisioGraphConnectorKind.Control : VisioGraphConnectorKind.Standard;
    }

    private static VisioSequenceParticipantKind MapParticipantKind(string kind, OfficeVisioVisualConversionReport report) {
        foreach (string declaredName in Enum.GetNames(typeof(VisioSequenceParticipantKind))) {
            if (string.Equals(declaredName, kind, StringComparison.OrdinalIgnoreCase)) {
                return (VisioSequenceParticipantKind)Enum.Parse(typeof(VisioSequenceParticipantKind), declaredName, ignoreCase: false);
            }
        }
        report.Warn($"Sequence participant kind '{kind}' was mapped to Visio's generic participant shape.");
        return VisioSequenceParticipantKind.Participant;
    }

    private static VisioSequenceMessageKind MapMessageKind(VisualArtifactInterchangeEdge edge) {
        if (Contains(edge.Kind, "Async")) return VisioSequenceMessageKind.Async;
        if (Contains(edge.Kind, "Event")) return VisioSequenceMessageKind.Event;
        if (Contains(edge.Kind, "Return") || string.Equals(edge.LineStyle, "Dashed", StringComparison.OrdinalIgnoreCase)) return VisioSequenceMessageKind.Return;
        return VisioSequenceMessageKind.Call;
    }

    private static VisioSide MapNoteSide(string? placement) =>
        string.Equals(placement, "LeftOf", StringComparison.OrdinalIgnoreCase) || string.Equals(placement, "Left", StringComparison.OrdinalIgnoreCase)
            ? VisioSide.Left
            : VisioSide.Right;

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

    private static void ReportArtifactAccessibilityFidelity(
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualConversionReport report) {
        if (!string.IsNullOrWhiteSpace(envelope.AccessibleName) ||
            !string.IsNullOrWhiteSpace(envelope.AccessibleDescription) ||
            !string.IsNullOrWhiteSpace(envelope.Language) ||
            envelope.IsDecorative) {
            report.Warn("Artifact accessibility and language semantics remain in the CFX envelope because the native Visio projection has no equivalent page-level contract.");
        }
    }

    private static bool Contains(string? value, string token) =>
        value != null && value.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;

    private static int ReadInt(IReadOnlyDictionary<string, string> metadata, string key, int fallback) =>
        metadata.TryGetValue(key, out string? value) && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : fallback;

    private static bool ReadBool(IReadOnlyDictionary<string, string> metadata, string key) =>
        metadata.TryGetValue(key, out string? value) && string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);

}
