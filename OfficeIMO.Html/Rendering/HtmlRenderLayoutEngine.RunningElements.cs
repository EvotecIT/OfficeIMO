using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlCssRunningStringAssignment CaptureRunningElement(
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
        return new HtmlCssRunningStringAssignment(
            HtmlCssRunningElementKeys.ForName(name),
            HtmlCssRunningElementParser.FormatSnapshotId(snapshotId),
            0D,
            orderOffset,
            GetDocumentOrder(element));
    }

    private static IReadOnlyList<HtmlCssRunningStringAssignment> NormalizeRunningElementAssignmentOrder(
        IEnumerable<HtmlCssRunningStringAssignment> assignments,
        double extent) {
        List<HtmlCssRunningStringAssignment> materialized = assignments.ToList();
        List<HtmlCssRunningStringAssignment> runningElements = materialized
            .Where(assignment => assignment.DocumentOrder.HasValue)
            .OrderBy(assignment => assignment.DocumentOrder)
            .ThenBy(assignment => assignment.OrderOffset)
            .ToList();
        if (runningElements.Count == 0) return materialized.OrderBy(assignment => assignment.OrderOffset).ToList();

        var logicalOffsets = new Dictionary<HtmlCssRunningStringAssignment, double>();
        double step = Math.Max(0.01D, extent) / runningElements.Count;
        for (int index = 0; index < runningElements.Count; index++) {
            logicalOffsets[runningElements[index]] = index * step;
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
