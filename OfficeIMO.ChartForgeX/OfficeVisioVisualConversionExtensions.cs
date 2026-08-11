using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.ChartForgeX;

/// <summary>Projects CFX semantic artifacts into native editable OfficeIMO.Visio diagrams.</summary>
public static class OfficeVisioVisualConversionExtensions {
    /// <summary>Projects a typed CFX artifact into a native editable Visio document.</summary>
    public static OfficeVisioVisualConversionResult ToOfficeVisio(
        this VisualArtifact artifact,
        OfficeVisioVisualOptions? options = null,
        VisualArtifactRenderOptions? renderOptions = null) {
        if (artifact == null) throw new ArgumentNullException(nameof(artifact));
        return artifact.ToInterchangeEnvelope(renderOptions).ToOfficeVisio(options);
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
        document.Title = string.IsNullOrWhiteSpace(validated.Title) ? null : validated.Title;
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
        if (envelope.Metadata.Count > 0) {
            report.Warn("Artifact-level metadata remains in the CFX envelope and is not duplicated into the native Visio graph page or document.");
        }

        document.GraphDiagram(options.PageName, builder => {
            ConfigureGraph(builder, envelope, options, report);
            builder.Import(nodes, edges, groups);
        });
    }

    private static void ConfigureGraph(
        VisioGraphDiagramBuilder builder,
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        (double width, double height) = ResolvePageSize(envelope, options);
        builder.PageSize(width, height).FitPageToGraph();
        builder.Layout(MapLayout(envelope.Layout));
        builder.Direction(MapDirection(envelope.Direction, report));
        if (envelope.Nodes.Any(node => node.X.HasValue || node.Y.HasValue)) {
            report.Warn("Native Visio layout was recomputed; prepared CFX pixel coordinates remain available in the semantic envelope.");
        }
        if (options.IncludeTitle && !string.IsNullOrWhiteSpace(envelope.Title)) {
            builder.Title(envelope.Title, UniqueTitleId(envelope));
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
            HyperlinkDescription = options.IncludeHyperlinks ? node.Tooltip : null
        };
        if (options.IncludeShapeData) {
            AddCommonShapeData(record.ShapeData, node.Kind, node.Status, node.GroupId, node.Metadata, report, "node '" + node.Id + "'");
            AddValue(record.ShapeData, "CFX.Icon", node.IconId);
            AddValue(record.ShapeData, "CFX.Symbol", node.Symbol);
            AddValue(record.ShapeData, "CFX.Badge", node.Badge);
            AddValue(record.ShapeData, "CFX.Color", node.Color);
            AddValue(record.ShapeData, "CFX.BackgroundColor", node.BackgroundColor);
            for (int index = 0; index < node.Details.Count; index++) {
                VisualArtifactInterchangeDetail detail = node.Details[index];
                AddValue(record.ShapeData, "Detail." + (index + 1).ToString(CultureInfo.InvariantCulture) + "." + detail.Label, detail.Value);
            }
            for (int index = 0; index < node.Ports.Count; index++) {
                VisualArtifactInterchangePort port = node.Ports[index];
                AddValue(record.ShapeData, "Port." + (index + 1).ToString(CultureInfo.InvariantCulture),
                    port.Id + "|" + port.Side + "|" + port.Offset.ToString("R", CultureInfo.InvariantCulture));
            }
        }
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
            HyperlinkDescription = options.IncludeHyperlinks ? edge.Tooltip : null
        };
        ApplyGraphEdgeDirection(record, edge.Direction, edge.Id, report);
        if (options.IncludeShapeData) {
            AddCommonShapeData(record.ShapeData, edge.Kind, edge.Status, null, edge.Metadata, report, "edge '" + edge.Id + "'");
            AddValue(record.ShapeData, "CFX.Direction", edge.Direction);
            AddValue(record.ShapeData, "CFX.LineStyle", edge.LineStyle);
            AddValue(record.ShapeData, "CFX.SourcePort", edge.SourcePortId ?? edge.SourcePort);
            AddValue(record.ShapeData, "CFX.TargetPort", edge.TargetPortId ?? edge.TargetPort);
            AddValue(record.ShapeData, "CFX.Color", edge.Color);
        }
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
                HyperlinkDescription = options.IncludeHyperlinks ? group.Tooltip : null
            };
            if (options.IncludeShapeData) {
                AddCommonShapeData(record.ShapeData, group.Kind, group.Status, null, group.Metadata, report, "group '" + group.Id + "'");
                AddValue(record.ShapeData, "CFX.Color", group.Color);
            }
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
            .ThenBy(node => node.Id, StringComparer.Ordinal)
            .ToList();
        List<VisualArtifactInterchangeEdge> messages = envelope.Edges.OrderBy(edge => edge.Order).ThenBy(edge => edge.Id, StringComparer.Ordinal).ToList();
        (double width, double height) = ResolveSequenceLayoutPageSize(envelope, options);
        bool includeTitle = options.IncludeTitle && !string.IsNullOrWhiteSpace(envelope.Title);
        var ids = new SequenceVisioIdMap(participants, messages, envelope.Annotations, includeTitle);

        ReportSequenceIdMappings(participants, messages, envelope.Annotations, ids, report);

        document.SequenceDiagram(options.PageName, builder => {
            builder.PageSize(width, height);
            if (ids.TitleId != null) {
                builder.Title(envelope.Title, ids.TitleId);
            }
            foreach (VisualArtifactInterchangeNode participant in participants) {
                builder.Participant(ids.Participant(participant.Id), participant.Label, MapParticipantKind(participant.Kind, report));
            }
            foreach (VisualArtifactInterchangeEdge message in messages) {
                VisioSequenceMessageKind kind = MapMessageKind(message);
                if (string.Equals(message.SourceId, message.TargetId, StringComparison.Ordinal)) {
                    builder.SelfMessage(ids.Participant(message.SourceId), message.Label ?? string.Empty, kind, ids.Message(message.Id));
                } else {
                    builder.Message(ids.Participant(message.SourceId), ids.Participant(message.TargetId), message.Label ?? string.Empty, kind, ids.Message(message.Id));
                }
            }
            AddSequenceActivations(builder, messages, ids);
            report.AnnotationCount = AddSequenceAnnotations(builder, envelope, messages.Count, ids, report);
        });

        VisioPage page = document.Pages[document.Pages.Count - 1];
        ApplySequenceParticipantData(page, participants, ids, options, report);
        ApplySequenceMessageData(page, messages, ids, options, report);
        ApplySequenceAnnotationData(page, envelope.Annotations, ids, options, report);
        if (options.UseNaturalPageSize) {
            page.Width = Math.Max(page.Width, width);
            page.Height = Math.Max(page.Height, height);
        } else {
            page.FitToContent(0.5D);
        }

        if (options.IncludeShapeData && envelope.Metadata.Count > 0) {
            report.Warn("Sequence-level metadata remains available in the CFX envelope; participant metadata is projected into native Visio Shape Data.");
        }
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
        foreach (KeyValuePair<string, Stack<int>> item in open) {
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
        int messageCount,
        SequenceVisioIdMap ids,
        OfficeVisioVisualConversionReport report) {
        int lastRow = Math.Max(0, messageCount - 1);
        int projected = 0;
        foreach (VisualArtifactInterchangeAnnotation annotation in envelope.Annotations) {
            int start = Clamp(annotation.StartIndex ?? 0, 0, lastRow);
            int end = Clamp(annotation.EndIndex ?? start, start, lastRow);
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
        return (Math.Max(4D, width), Math.Max(3D, height));
    }

    private static (double Width, double Height) ResolveSequenceLayoutPageSize(VisualArtifactInterchangeEnvelope envelope, OfficeVisioVisualOptions options) {
        if (!options.UseNaturalPageSize) return (1D, 1D);
        double width = envelope.Width.HasValue ? envelope.Width.Value / options.PixelsPerInch : 11D;
        double height = envelope.Height.HasValue ? envelope.Height.Value / options.PixelsPerInch : 8.5D;
        return (Math.Max(11D, width), Math.Max(8.5D, height));
    }

    private static void ApplySequenceParticipantData(
        VisioPage page,
        IEnumerable<VisualArtifactInterchangeNode> participants,
        SequenceVisioIdMap ids,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        foreach (VisualArtifactInterchangeNode participant in participants) {
            VisioShape shape = page.Shapes.Single(item => string.Equals(item.Id, ids.Participant(participant.Id), StringComparison.Ordinal));
            if (options.IncludeShapeData) {
                var data = new Dictionary<string, string?>(StringComparer.Ordinal);
                AddValue(data, "CFX.Id", participant.Id);
                AddCommonShapeData(data, participant.Kind, participant.Status, participant.GroupId, participant.Metadata, report, "sequence participant '" + participant.Id + "'");
                AddValue(data, "CFX.Icon", participant.IconId);
                AddValue(data, "CFX.Symbol", participant.Symbol);
                AddValue(data, "CFX.Badge", participant.Badge);
                AddValue(data, "CFX.Color", participant.Color);
                AddValue(data, "CFX.BackgroundColor", participant.BackgroundColor);
                for (int index = 0; index < participant.Details.Count; index++) {
                    VisualArtifactInterchangeDetail detail = participant.Details[index];
                    AddValue(data, "Detail." + (index + 1).ToString(CultureInfo.InvariantCulture) + "." + detail.Label, detail.Value);
                }
                for (int index = 0; index < participant.Ports.Count; index++) {
                    VisualArtifactInterchangePort port = participant.Ports[index];
                    AddValue(data, "Port." + (index + 1).ToString(CultureInfo.InvariantCulture),
                        port.Id + "|" + port.Side + "|" + port.Offset.ToString("R", CultureInfo.InvariantCulture));
                }
                foreach (KeyValuePair<string, string?> item in data) shape.SetShapeData(item.Key, item.Value);
            }
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
            if (options.IncludeShapeData) {
                var data = new Dictionary<string, string?>(StringComparer.Ordinal);
                AddValue(data, "CFX.Id", message.Id);
                AddCommonShapeData(data, message.Kind, message.Status, null, message.Metadata, report, "sequence message '" + message.Id + "'");
                AddValue(data, "CFX.Direction", message.Direction);
                AddValue(data, "CFX.LineStyle", message.LineStyle);
                AddValue(data, "CFX.SourcePort", message.SourcePortId ?? message.SourcePort);
                AddValue(data, "CFX.TargetPort", message.TargetPortId ?? message.TargetPort);
                AddValue(data, "CFX.Color", message.Color);
                foreach (KeyValuePair<string, string?> item in data) connector.SetShapeData(item.Key, item.Value);
            }
            if (options.IncludeHyperlinks && !string.IsNullOrWhiteSpace(message.Href)) {
                connector.AddHyperlink(message.Href!, message.Tooltip);
            }
        }
    }

    private static void ApplySequenceAnnotationData(
        VisioPage page,
        IEnumerable<VisualArtifactInterchangeAnnotation> annotations,
        SequenceVisioIdMap ids,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        foreach (VisualArtifactInterchangeAnnotation annotation in annotations) {
            VisioShape? shape = page.Shapes.FirstOrDefault(item => string.Equals(item.Id, ids.Annotation(annotation.Id), StringComparison.Ordinal));
            if (shape == null || !options.IncludeShapeData) continue;
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

    private static VisioGraphLayout MapLayout(string layout) {
        if (string.Equals(layout, "Dense", StringComparison.OrdinalIgnoreCase) || string.Equals(layout, "Grid", StringComparison.OrdinalIgnoreCase)) return VisioGraphLayout.Grid;
        if (string.Equals(layout, "Force", StringComparison.OrdinalIgnoreCase) || string.Equals(layout, "Radial", StringComparison.OrdinalIgnoreCase)) return VisioGraphLayout.Radial;
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
        if (Enum.TryParse(kind, true, out VisioSequenceParticipantKind parsed)) return parsed;
        if (Contains(kind, "Queue") || Contains(kind, "Collection")) {
            report.Warn($"Sequence participant kind '{kind}' was mapped to Visio's generic participant shape.");
        }
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

    private static void AddCommonShapeData(
        IDictionary<string, string?> target,
        string kind,
        string? status,
        string? groupId,
        IEnumerable<KeyValuePair<string, string>> metadata,
        OfficeVisioVisualConversionReport report,
        string context) {
        AddValue(target, "CFX.Kind", kind);
        AddValue(target, "CFX.Status", status);
        AddValue(target, "CFX.GroupId", groupId);
        foreach (KeyValuePair<string, string> item in metadata.OrderBy(pair => pair.Key, StringComparer.Ordinal)) {
            string requested = "Metadata." + item.Key;
            string resolved = requested;
            int suffix = 2;
            while (target.Keys.Any(key => string.Equals(key, resolved, StringComparison.OrdinalIgnoreCase))) {
                resolved = requested + " [" + suffix.ToString(CultureInfo.InvariantCulture) + "]";
                suffix++;
            }
            if (!string.Equals(requested, resolved, StringComparison.Ordinal)) {
                report.Warn($"Metadata key '{item.Key}' on {context} was projected as '{resolved}' because Visio Shape Data names are case-insensitive.");
            }
            AddValue(target, resolved, item.Value);
        }
    }

    private static void AddValue(IDictionary<string, string?> target, string key, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) target[key] = value;
    }

    private static string CombineLabel(string primary, string? secondary) =>
        string.IsNullOrWhiteSpace(secondary) ? primary : primary + Environment.NewLine + secondary;

    private static string? CombineEdgeLabel(VisualArtifactInterchangeEdge edge) {
        string[] labels = new[] { edge.Label, edge.SecondaryLabel, edge.TertiaryLabel }
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value!)
            .ToArray();
        return labels.Length == 0 ? null : string.Join(" | ", labels);
    }

    private static bool Contains(string? value, string token) =>
        value != null && value.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;

    private static int ReadInt(IReadOnlyDictionary<string, string> metadata, string key, int fallback) =>
        metadata.TryGetValue(key, out string? value) && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : fallback;

    private static bool ReadBool(IReadOnlyDictionary<string, string> metadata, string key) =>
        metadata.TryGetValue(key, out string? value) && string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);

    private static int Clamp(int value, int minimum, int maximum) => value < minimum ? minimum : value > maximum ? maximum : value;
}
