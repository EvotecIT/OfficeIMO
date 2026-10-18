using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Rebuilds an ordinary owner when its shared fragment planner moves a sibling keep.
    /// Each accepted child-entry gap is applied before that child's descendants establish their page bands.
    /// </summary>
    private HtmlRenderFlowBlock LayoutElement(IElement element, double containingWidth,
        HtmlRenderBoxStyle style, HtmlRenderBoxStyle parentStyle, int depth,
        IElement? continuationTarget = null, int continuationLogicalCharacters = 0,
        PagedFloatBoundary? pageBoundary = null, IReadOnlyList<HtmlFloatExclusion>? inheritedFloats = null,
        ICollection<HtmlFloatExclusion>? emittedFloats = null, InlineFloatContext? surroundingFloats = null) {
        ConstrainedSourceProjection? context = _constrainedSourceProjection;
        double entryGap = 0D;
        if (context != null && context.AtomicDepth == 0) {
            foreach (ConstrainedInteriorRelayoutFrame ancestor in context.InteriorRelayoutFrames) {
                CheckCancellation();
                ChargeLayoutOperation("constrained planned entry lookup");
                if (!ancestor.EntryGaps.TryGetValue(element, out ConstrainedScheduledEntry? entry) || entry.Applied) continue;
                entry.Applied = true;
                entryGap = entry.Gap;
                break;
            }
        }
        PagedFloatBoundary? physical = pageBoundary?.Shift(entryGap);
        var frame = context == null ? null : new ConstrainedInteriorRelayoutFrame();
        if (frame != null) context!.InteriorRelayoutFrames.Push(frame);
        try {
            while (true) {
                CheckCancellation();
                if (frame != null) {
                    foreach (ConstrainedScheduledEntry entry in frame.EntryGaps.Values) {
                        CheckCancellation();
                        ChargeLayoutOperation("constrained child-entry replay");
                        entry.Applied = false;
                    }
                }
                HtmlRenderFlowBlock result = LayoutElementCore(element, containingWidth, style, parentStyle, depth,
                    continuationTarget, continuationLogicalCharacters, physical, inheritedFloats, emittedFloats,
                    surroundingFloats, out ConstrainedEntryRelayoutState? checkpoint, out ConstrainedInteriorGap? gap);
                if (gap == null || context!.FailedReason != null) {
                    if (frame != null) ChargeLayoutOperations(frame.EntryGaps.Count, "constrained child-entry progress");
                    if (frame != null && frame.EntryGaps.Values.Any(entry => !entry.Applied))
                        context!.FailedReason = "constrained-interior-entry-owner-not-rebuilt";
                    if (entryGap > 0.0001D) {
                        result = result.AdjustLeadingFlowSpace(-entryGap);
                        result.PagedSourceEntryGap += entryGap;
                        context!.Changed = true;
                        context.NeedsProjection = true;
                    }
                    return result;
                }
                ChargeLayoutOperation("constrained interior owner relayout");
                if (checkpoint == null || frame == null || frame.EntryGaps.ContainsKey(gap.Element)) {
                    context.FailedReason = "constrained-interior-entry-without-source-progress";
                    return result;
                }
                frame.EntryGaps.Add(gap.Element, new ConstrainedScheduledEntry(gap.Gap));
                RestoreConstrainedEntryState(checkpoint, context);
                style = checkpoint.Style.Clone();
            }
        } finally {
            if (frame != null) context!.InteriorRelayoutFrames.Pop();
        }
    }

    private sealed record ConstrainedInteriorGap(IElement Element, double Gap);

    private sealed class ConstrainedInteriorRelayoutFrame {
        internal Dictionary<IElement, ConstrainedScheduledEntry> EntryGaps { get; } = new();
    }

    private sealed class ConstrainedScheduledEntry {
        internal ConstrainedScheduledEntry(double gap) => Gap = gap;
        internal double Gap { get; }
        internal bool Applied { get; set; }
    }
}
