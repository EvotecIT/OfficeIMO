namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static bool IsForcedPageBreak(HtmlPageBreakTarget target) =>
        target is HtmlPageBreakTarget.Page or HtmlPageBreakTarget.Left or HtmlPageBreakTarget.Right;

    private HtmlPageBreakTarget ApplicableForcedFlowBreak(HtmlPageBreakTarget target) =>
        target == HtmlPageBreakTarget.Column && _columnContentDepth == 0 ? HtmlPageBreakTarget.None : target;

    private static bool IsForcedBreakForScope(HtmlPageBreakTarget target, HtmlFragmentationScope scope) =>
        IsForcedPageBreak(target) || target == HtmlPageBreakTarget.Column && (scope & HtmlFragmentationScope.Column) != 0;

    /// <summary>
    /// Preserves authored sibling boundaries through normal-flow wrappers. Edge
    /// values propagate to the enclosing boundary, including through decorations;
    /// internal empty siblings retain distinct identities even at equal offsets.
    /// </summary>
    private void AppendChildForcedBreaks(IReadOnlyList<HtmlRenderFlowBlock> children, int index,
        double start, ICollection<HtmlRenderForcedBreak> result, object[] boundaries) {
        HtmlRenderFlowBlock child = children[index];
        HtmlRenderBreakEdge before = index == 0 ? HtmlRenderBreakEdge.Leading : HtmlRenderBreakEdge.Interior;
        HtmlRenderBreakEdge after = index == children.Count - 1 ? HtmlRenderBreakEdge.Trailing : HtmlRenderBreakEdge.Interior;
        if (child.BreakBefore != HtmlPageBreakTarget.None)
            result.Add(new HtmlRenderForcedBreak(start, child.BreakBefore, boundary: boundaries[index], edge: before));
        foreach (HtmlRenderForcedBreak item in child.ForcedBreaks) {
            CheckCancellation();
            ChargeLayoutOperation("forced flow boundary propagation");
            result.Add(item.Edge switch {
                HtmlRenderBreakEdge.Leading => item.AtBoundary(start, boundaries[index], before),
                HtmlRenderBreakEdge.Trailing => item.AtBoundary(start + child.NormalFlowAdvance, boundaries[index + 1], after),
                _ => item.Translate(start)
            });
        }
        if (child.BreakAfter != HtmlPageBreakTarget.None)
            result.Add(new HtmlRenderForcedBreak(start + child.NormalFlowAdvance, child.BreakAfter,
                boundary: boundaries[index + 1], edge: after, isAfter: true));
    }

    private object[] CreateFlowBoundaries(IReadOnlyList<HtmlRenderFlowBlock> children) {
        if (!children.Any(child => child.BreakBefore != HtmlPageBreakTarget.None
            || child.BreakAfter != HtmlPageBreakTarget.None || child.ForcedBreaks.Count > 0)) return Array.Empty<object>();
        int count = children.Count;
        CheckCancellation();
        ChargeLayoutOperations((long)count + 1L, "forced sibling boundaries");
        var boundaries = new object[count + 1];
        for (int index = 0; index < boundaries.Length; index++) boundaries[index] = new object();
        return boundaries;
    }
}
