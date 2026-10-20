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
        double entryGap = TakeConstrainedScheduledEntryGap(element);
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
                    return ApplyConstrainedScheduledEntryGap(result, entryGap);
                }
                ChargeLayoutOperation("constrained interior owner relayout");
                if (checkpoint == null || frame == null || frame.EntryGaps.ContainsKey(gap.SourceEntry)) {
                    context.FailedReason = "constrained-interior-entry-without-source-progress";
                    return result;
                }
                frame.EntryGaps.Add(gap.SourceEntry, new ConstrainedScheduledEntry(gap.Gap));
                RestoreConstrainedEntryState(checkpoint, context);
                style = checkpoint.Style.Clone();
            }
        } finally {
            if (frame != null) context!.InteriorRelayoutFrames.Pop();
        }
    }

    /// <summary>Consumes one planned source entry by original DOM identity, including anonymous text runs.</summary>
    private double TakeConstrainedScheduledEntryGap(INode sourceEntry) {
        ConstrainedSourceProjection? context = _constrainedSourceProjection;
        if (context == null || context.AtomicDepth != 0) return 0D;
        foreach (ConstrainedInteriorRelayoutFrame ancestor in context.InteriorRelayoutFrames) {
            CheckCancellation();
            ChargeLayoutOperation("constrained planned entry lookup");
            if (!ancestor.EntryGaps.TryGetValue(sourceEntry, out ConstrainedScheduledEntry? entry) || entry.Applied) continue;
            entry.Applied = true;
            return entry.Gap;
        }
        return 0D;
    }

    /// <summary>Represents a consumed entry before its source without adding to authored content height.</summary>
    private HtmlRenderFlowBlock ApplyConstrainedScheduledEntryGap(HtmlRenderFlowBlock block, double entryGap) {
        if (entryGap <= 0.0001D) return block;
        HtmlRenderFlowBlock result = block.AdjustLeadingFlowSpace(-entryGap);
        result.PagedSourceEntryGap += entryGap;
        ConstrainedSourceProjection context = _constrainedSourceProjection!;
        context.Changed = true;
        context.NeedsProjection = true;
        return result;
    }

    private sealed record ConstrainedInteriorGap(INode SourceEntry, double Gap);

    private sealed class ConstrainedInteriorRelayoutFrame {
        internal Dictionary<INode, ConstrainedScheduledEntry> EntryGaps { get; } = new();
    }

    private sealed class ConstrainedScheduledEntry {
        internal ConstrainedScheduledEntry(double gap) => Gap = gap;
        internal double Gap { get; }
        internal bool Applied { get; set; }
    }
}
