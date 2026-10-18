using AngleSharp.Dom;

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
        double beforeEntryGap = child.PagedSourceEntryGap;
        double afterEntryGap = index + 1 < children.Count ? children[index + 1].PagedSourceEntryGap : 0D;
        if (child.BreakBefore != HtmlPageBreakTarget.None)
            result.Add(new HtmlRenderForcedBreak(start + beforeEntryGap, child.BreakBefore,
                boundary: boundaries[index], edge: before, projected: beforeEntryGap > 0.0001D));
        foreach (HtmlRenderForcedBreak item in child.ForcedBreaks) {
            CheckCancellation();
            ChargeLayoutOperation("forced flow boundary propagation");
            result.Add(item.Edge switch {
                HtmlRenderBreakEdge.Leading => BindEntry(item, start + beforeEntryGap, boundaries[index], before, beforeEntryGap),
                HtmlRenderBreakEdge.Trailing => BindEntry(item, start + child.NormalFlowAdvance + afterEntryGap,
                    boundaries[index + 1], after, afterEntryGap),
                _ => item.Translate(start)
            });
        }
        if (child.BreakAfter != HtmlPageBreakTarget.None)
            result.Add(new HtmlRenderForcedBreak(start + child.NormalFlowAdvance + afterEntryGap, child.BreakAfter,
                boundary: boundaries[index + 1], edge: after, isAfter: true, projected: afterEntryGap > 0.0001D,
                advancesFollowingFlow: child.NormalFlowAdvance > 0.0001D || child.Visuals.Any(HasOrdinaryFragmentSource)));

        static HtmlRenderForcedBreak BindEntry(HtmlRenderForcedBreak item, double offset, object boundary,
            HtmlRenderBreakEdge edge, double entryGap) {
            HtmlRenderForcedBreak rebound = item.AtBoundary(offset, boundary, edge);
            return entryGap > 0.0001D && IsForcedPageBreak(item.Target) ? rebound.ProjectAt(offset) : rebound;
        }
    }

    private object[] CreateFlowBoundaries(IReadOnlyList<HtmlRenderFlowBlock> children) {
        if (!children.Any(child => child.BreakBefore != HtmlPageBreakTarget.None
            || child.BreakAfter != HtmlPageBreakTarget.None || child.ForcedBreaks.Count > 0)) return Array.Empty<object>();
        int count = children.Count;
        CheckCancellation();
        ChargeLayoutOperations((long)count + 1L, "forced sibling boundaries");
        var boundaries = new object[count + 1];
        for (int index = 0; index < boundaries.Length; index++) {
            HtmlRenderFlowBlock? preceding = index > 0 ? children[index - 1] : null;
            boundaries[index] = new OrdinaryFlowBoundary(index < count ? children[index].OwnerElement : null,
                preceding != null && preceding.PagedPaintExtent > preceding.NormalFlowAdvance + 0.0001D);
        }
        return boundaries;
    }

    /// <summary>Retains the real following child when a source boundary crosses its preceding sibling's visible overflow.</summary>
    private sealed class OrdinaryFlowBoundary(IElement? followingElement, bool precedingPaintOverhang) {
        internal IElement? FollowingElement { get; } = followingElement;
        internal bool PrecedingPaintOverhang { get; } = precedingPaintOverhang;
    }
}
