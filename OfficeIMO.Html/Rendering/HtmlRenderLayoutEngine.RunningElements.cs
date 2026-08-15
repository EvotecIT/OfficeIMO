using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<HtmlCssRunningStringAssignment> CaptureRunningElement(
        IElement element,
        string name,
        double containingWidth,
        HtmlRenderBoxStyle style,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        double orderOffset = 0D) {
        HtmlRenderBoxStyle captureStyle = style.Clone();
        captureStyle.Position = "static";
        captureStyle.ZIndex = "auto";
        HtmlRenderFlowBlock snapshot = LayoutElement(element, containingWidth, captureStyle, parentStyle, depth);
        int snapshotId = ++_nextRunningElementSnapshotId;
        _runningElementSnapshots[snapshotId] = new HtmlCssRunningElementSnapshot(snapshot, element, parentStyle, depth);
        int documentOrder = GetDocumentOrder(element);
        var assignments = new List<HtmlCssRunningStringAssignment>(snapshot.RunningStringAssignments.Count + 1) {
            new HtmlCssRunningStringAssignment(
                HtmlCssRunningElementKeys.ForName(name),
                HtmlCssRunningElementParser.FormatSnapshotId(snapshotId),
                0D,
                orderOffset,
                documentOrder)
        };
        assignments.AddRange(snapshot.RunningStringAssignments.Select(assignment =>
            new HtmlCssRunningStringAssignment(
                assignment.Name,
                assignment.Value,
                0D,
                orderOffset,
                documentOrder)));
        return assignments.AsReadOnly();
    }

    private static IReadOnlyList<HtmlCssRunningStringAssignment> NormalizeRunningElementAssignmentOrder(
        IEnumerable<HtmlCssRunningStringAssignment> assignments,
        double extent) {
        List<HtmlCssRunningStringAssignment> materialized = assignments.ToList();
        List<IGrouping<int, HtmlCssRunningStringAssignment>> documentOrderGroups = materialized
            .Where(assignment => assignment.DocumentOrder.HasValue)
            .GroupBy(assignment => assignment.DocumentOrder!.Value)
            .OrderBy(group => group.Key)
            .ToList();
        if (documentOrderGroups.Count == 0) return materialized.OrderBy(assignment => assignment.OrderOffset).ToList();

        var logicalOffsets = new Dictionary<HtmlCssRunningStringAssignment, double>();
        double step = Math.Max(0.01D, extent) / documentOrderGroups.Count;
        for (int index = 0; index < documentOrderGroups.Count; index++) {
            foreach (HtmlCssRunningStringAssignment assignment in documentOrderGroups[index]) {
                logicalOffsets[assignment] = index * step;
            }
        }

        return materialized
            .Select(assignment => logicalOffsets.TryGetValue(assignment, out double orderOffset)
                ? assignment.Place(assignment.Offset, orderOffset)
                : assignment)
            .OrderBy(assignment => assignment.OrderOffset)
            .ToList();
    }

    private static IReadOnlyList<HtmlCssRunningStringAssignment> PlaceDirectRunningElementAssignments(
        IEnumerable<HtmlCssRunningStringAssignment> assignments,
        IEnumerable<RunningElementFlowAnchor> anchors,
        double startOffset,
        double endOffset) {
        List<RunningElementFlowAnchor> orderedAnchors = anchors
            .OrderBy(anchor => anchor.SourceIndex)
            .ToList();
        if (orderedAnchors.Count == 0) {
            return assignments
                .Select(assignment => assignment.Place(startOffset, assignment.OrderOffset))
                .ToList();
        }

        var placed = new List<HtmlCssRunningStringAssignment>();
        foreach (HtmlCssRunningStringAssignment assignment in assignments) {
            RunningElementFlowAnchor? following = orderedAnchors
                .FirstOrDefault(anchor => anchor.SourceIndex > assignment.OrderOffset);
            double offset = following?.Offset ?? endOffset;
            placed.Add(assignment.Place(offset, assignment.OrderOffset));
        }
        return placed;
    }

    private sealed class RunningElementFlowAnchor {
        internal RunningElementFlowAnchor(int sourceIndex, double offset) {
            SourceIndex = sourceIndex;
            Offset = offset;
        }

        internal int SourceIndex { get; }
        internal double Offset { get; }
    }
}
