using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.ChartForgeX;

public static partial class OfficeVisioVisualConversionExtensions {
    private sealed class SequenceActivationChange {
        public SequenceActivationChange(int row, int sourceOrder, int ordinal, string participant, bool active, string entityId) {
            Row = row;
            SourceOrder = sourceOrder;
            Ordinal = ordinal;
            Participant = participant;
            Active = active;
            EntityId = entityId;
        }

        public int Row { get; }
        public int SourceOrder { get; }
        public int Ordinal { get; }
        public string Participant { get; }
        public bool Active { get; }
        public string EntityId { get; }
        public List<VisualArtifactInterchangeAnnotation> Annotations { get; } = new();
    }

    private sealed class SequenceActivationOpen {
        public SequenceActivationOpen(int row, IEnumerable<VisualArtifactInterchangeAnnotation> annotations) {
            Row = row;
            Annotations = annotations.ToList();
        }

        public int Row { get; }
        public IReadOnlyList<VisualArtifactInterchangeAnnotation> Annotations { get; }
    }

    private sealed class SequenceActivationProjection {
        public SequenceActivationProjection(string shapeId, string participant, int start, int end, IReadOnlyList<VisualArtifactInterchangeAnnotation> annotations) {
            ShapeId = shapeId;
            Participant = participant;
            Start = start;
            End = end;
            Annotations = annotations;
        }

        public string ShapeId { get; }
        public string Participant { get; }
        public int Start { get; }
        public int End { get; }
        public IReadOnlyList<VisualArtifactInterchangeAnnotation> Annotations { get; }
    }

    private sealed class SequenceBlockProjection {
        public SequenceBlockProjection(VisualArtifactInterchangeAnnotation annotation, string shapeId, int depth, IReadOnlyList<string> participantIds) {
            Annotation = annotation;
            ShapeId = shapeId;
            Depth = depth;
            ParticipantIds = participantIds;
        }

        public VisualArtifactInterchangeAnnotation Annotation { get; }
        public string ShapeId { get; }
        public int Depth { get; }
        public IReadOnlyList<string> ParticipantIds { get; }
    }

    private static SequenceActivationProjection CreateActivationProjection(
        SequenceVisioIdMap ids,
        string participant,
        int start,
        int end,
        IEnumerable<VisualArtifactInterchangeAnnotation> startAnnotations,
        IEnumerable<VisualArtifactInterchangeAnnotation> endAnnotations) {
        List<VisualArtifactInterchangeAnnotation> annotations = startAnnotations
            .Concat(endAnnotations)
            .GroupBy(annotation => annotation.Id, StringComparer.Ordinal)
            .Select(group => group.First())
            .ToList();
        string shapeId = annotations.Count == 0 ? ids.Activation() : ids.Annotation(annotations[0].Id);
        return new SequenceActivationProjection(shapeId, participant, start, end, annotations);
    }

    private static int AddSequenceFragments(
        VisioSequenceDiagramBuilder builder,
        VisualArtifactInterchangeEnvelope envelope,
        SequenceVisioIdMap ids,
        OfficeVisioVisualConversionReport report) {
        string[] allParticipants = envelope.Nodes.Select(node => ids.Participant(node.Id)).ToArray();
        var projectedBlocks = new List<SequenceBlockProjection>();
        int projected = 0;
        foreach (VisualArtifactInterchangeAnnotation block in envelope.Annotations
                     .Where(annotation => annotation.Role == VisualArtifactInterchangeAnnotationRole.SequenceBlock)
                     .OrderBy(annotation => annotation.StartIndex)
                     .ThenByDescending(annotation => annotation.EndIndex)
                     .ThenBy(annotation => annotation.Id, StringComparer.Ordinal)) {
            if (block.Sequence!.IsEmpty) {
                report.Warn(OfficeVisioVisualDiagnosticCode.AnnotationNotProjected, OfficeVisioVisualEntityKind.Annotation, block.Id, "emptyBlock",
                    $"Empty sequence block '{block.Id}' remains in the CFX envelope because a native Visio fragment requires a drawable row span.");
                continue;
            }

            int start = block.StartIndex!.Value;
            int end = block.EndIndex!.Value;
            SequenceBlockProjection? parent = projectedBlocks
                .Where(candidate => candidate.Annotation.StartIndex <= start && candidate.Annotation.EndIndex >= end)
                .OrderByDescending(candidate => candidate.Depth)
                .ThenBy(candidate => candidate.Annotation.EndIndex!.Value - candidate.Annotation.StartIndex!.Value)
                .FirstOrDefault();
            string nativeId = ids.Annotation(block.Id);
            IReadOnlyList<string> participants = block.TargetIds.Count == 0
                ? parent?.ParticipantIds ?? allParticipants
                : block.TargetIds.Select(ids.Participant).ToArray();
            string label = CombineLabel(block.Sequence.BlockKind!.Value.ToString(), block.Text);
            if (parent == null) {
                builder.Fragment(label, start, end, participants, nativeId);
            } else {
                builder.NestedFragment(parent.ShapeId, label, start, end, participants, nativeId);
            }
            projectedBlocks.Add(new SequenceBlockProjection(block, nativeId, (parent?.Depth ?? -1) + 1, participants));
            projected++;
        }

        foreach (VisualArtifactInterchangeAnnotation branch in envelope.Annotations
                     .Where(annotation => annotation.Role == VisualArtifactInterchangeAnnotationRole.SequenceBranch)
                     .OrderBy(annotation => annotation.StartIndex)
                     .ThenBy(annotation => annotation.Sequence!.Depth)
                     .ThenBy(annotation => annotation.Id, StringComparer.Ordinal)) {
            int start = branch.StartIndex!.Value;
            int end = branch.EndIndex ?? start;
            SequenceBlockProjection? parent = projectedBlocks
                .Where(candidate => candidate.Annotation.Sequence!.BlockKind == branch.Sequence!.ParentBlockKind &&
                                    candidate.Annotation.StartIndex <= start && candidate.Annotation.EndIndex >= end)
                .OrderBy(candidate => candidate.Depth == branch.Sequence!.Depth ? 0 : 1)
                .ThenByDescending(candidate => candidate.Depth)
                .ThenBy(candidate => candidate.Annotation.EndIndex!.Value - candidate.Annotation.StartIndex!.Value)
                .FirstOrDefault();
            if (parent == null) {
                report.Warn(OfficeVisioVisualDiagnosticCode.AnnotationNotProjected, OfficeVisioVisualEntityKind.Annotation, branch.Id, "branchParent",
                    $"Sequence branch '{branch.Id}' has no matching projected parent block and remains in the CFX envelope.");
                continue;
            }

            string nativeId = ids.Annotation(branch.Id);
            string label = CombineLabel(branch.Sequence!.BranchKind!, branch.Text);
            bool primary = string.Equals(branch.Sequence.BranchKind, "Primary", StringComparison.OrdinalIgnoreCase);
            if (primary || start <= parent.Annotation.StartIndex) {
                builder.FragmentGuard(parent.ShapeId, label, Math.Max(start, parent.Annotation.StartIndex!.Value), nativeId);
                if (!primary && start <= parent.Annotation.StartIndex) {
                    report.Warn(OfficeVisioVisualDiagnosticCode.SemanticLoss, OfficeVisioVisualEntityKind.Annotation, branch.Id, "branchDivider",
                        $"Sequence branch '{branch.Id}' was retained as a guard without a divider because its boundary is not after the parent fragment start.");
                }
            } else {
                builder.FragmentPartition(parent.ShapeId, label, start, nativeId);
            }
            projected++;
        }
        return projected;
    }

    private static void ApplySequenceAnnotationShapeData(
        VisioShape shape,
        VisualArtifactInterchangeAnnotation annotation,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        shape.Data["CFX.Id"] = annotation.Id;
        if (!options.IncludeShapeData) return;
        var data = new Dictionary<string, string?>(StringComparer.Ordinal);
        AddValue(data, "CFX.Id", annotation.Id);
        AddCommonShapeData(data, annotation.Kind, null, null, annotation.Extensions, report, "sequence annotation '" + annotation.Id + "'");
        AddValue(data, "CFX.Role", annotation.Role.ToString());
        AddValue(data, "CFX.Placement", annotation.Placement);
        AddValue(data, "CFX.TargetIds", annotation.TargetIds.Count == 0 ? null : string.Join(",", annotation.TargetIds));
        AddValue(data, "CFX.StartIndex", annotation.StartIndex?.ToString(CultureInfo.InvariantCulture));
        AddValue(data, "CFX.EndIndex", annotation.EndIndex?.ToString(CultureInfo.InvariantCulture));
        AddValue(data, "CFX.SequenceActivationState", annotation.Sequence?.ActivationState?.ToString(CultureInfo.InvariantCulture));
        AddValue(data, "CFX.SequenceNotePlacement", annotation.Sequence?.NotePlacement?.ToString());
        AddValue(data, "CFX.SequenceBlockKind", annotation.Sequence?.BlockKind?.ToString());
        AddValue(data, "CFX.SequenceParentBlockKind", annotation.Sequence?.ParentBlockKind?.ToString());
        AddValue(data, "CFX.SequenceBranchKind", annotation.Sequence?.BranchKind);
        AddValue(data, "CFX.SequenceDepth", annotation.Sequence?.Depth.ToString(CultureInfo.InvariantCulture));
        AddValue(data, "CFX.SequenceIsEmpty", annotation.Sequence?.IsEmpty.ToString(CultureInfo.InvariantCulture));
        foreach (KeyValuePair<string, string?> item in data) shape.SetShapeData(item.Key, item.Value);
    }

    private static void ApplySequenceActivationShapeData(
        VisioShape shape,
        SequenceActivationProjection projection,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        if (projection.Annotations.Count == 0) return;
        VisualArtifactInterchangeAnnotation first = projection.Annotations[0];
        ApplySequenceAnnotationShapeData(shape, first, options, report);
        string eventIds = string.Join(",", projection.Annotations.Select(annotation => annotation.Id));
        shape.Data["CFX.ActivationEventIds"] = eventIds;
        if (!options.IncludeShapeData) return;
        shape.SetShapeData("CFX.ActivationEventIds", eventIds);
        for (int index = 0; index < projection.Annotations.Count; index++) {
            VisualArtifactInterchangeAnnotation annotation = projection.Annotations[index];
            string prefix = "CFX.ActivationEvent." + (index + 1).ToString(CultureInfo.InvariantCulture) + ".";
            shape.SetShapeData(prefix + "Id", annotation.Id);
            shape.SetShapeData(prefix + "Kind", annotation.Kind);
            shape.SetShapeData(prefix + "State", annotation.Sequence!.ActivationState!.Value.ToString(CultureInfo.InvariantCulture));
            shape.SetShapeData(prefix + "StepIndex", annotation.StartIndex!.Value.ToString(CultureInfo.InvariantCulture));
            foreach (KeyValuePair<string, string> extension in annotation.Extensions) {
                shape.SetShapeData(prefix + "Extension." + extension.Key, extension.Value);
            }
        }
    }
}
