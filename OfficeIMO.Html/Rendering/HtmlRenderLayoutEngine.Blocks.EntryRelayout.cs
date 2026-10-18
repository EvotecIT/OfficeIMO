using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Retains the ordinary owner state needed if its entry keep moves already projected descendants.</summary>
    private ConstrainedEntryRelayoutState? CaptureConstrainedEntryRelayout(IElement element,
        HtmlRenderBoxStyle style, ConstrainedProjectionScope scope, PagedFloatBoundary? boundary,
        bool usesBlockFormatting) {
        ConstrainedSourceProjection? context = _constrainedSourceProjection;
        if (!scope.CanProject || !boundary.HasValue || context == null) return null;
        bool siblingKeep = usesBlockFormatting && HasConstrainedSiblingPageKeep(element);
        if (!context.TargetAncestors.Contains(element) && !(style.AvoidBreakInside && usesBlockFormatting)
            && !siblingKeep) return null;
        double available = boundary.Value.RemainingHeight
            - style.MarginTop - style.BorderTopWidth - style.PaddingTop;
        if (available <= 0.0001D)
            available += (Math.Floor(-available / context.PageHeight) + 1D) * context.PageHeight;
        if (available >= context.PageHeight - 0.0001D && !siblingKeep) return null;
        CheckCancellation();
        ChargeLayoutOperations(1L + _normalFlowPlacements.Count + _layoutStyles.Count
            + _inlineFloatOverhangs.Count + _semanticNodeIds.Count,
            "constrained entry relayout state");
        return new ConstrainedEntryRelayoutState(style.Clone(), _normalFlowPlacements.ToArray(),
            _layoutStyles.ToArray(), _inlineFloatOverhangs.ToArray(), _semanticNodeIds.ToArray(),
            _nextSemanticNodeId, _nextLogicalTextOrder, context.Changed, context.NeedsProjection);
    }

    /// <summary>Only ordinary sibling page keeps can request an interior child-entry rebuild.</summary>
    private bool HasConstrainedSiblingPageKeep(IElement element) {
        foreach (IElement child in element.Children) {
            CheckCancellation();
            ChargeLayoutOperation("constrained sibling keep ownership");
            if (!_layoutStyles.TryGetValue(child, out HtmlRenderBoxStyle? style)) continue;
            if (((style.AvoidBreakBeforeScope | style.AvoidBreakAfterScope) & HtmlFragmentationScope.Page) != 0)
                return true;
            if (style.Display == "contents" && HasConstrainedSiblingPageKeep(child)) return true;
        }
        return false;
    }

    /// <summary>
    /// Rebuilds nested ordinary sources after their ancestor moves to a fresh page.
    /// A discarded descendant gap is never reused at the ancestor's new origin.
    /// </summary>
    private bool TryRelayoutConstrainedEntry(IElement element, double containingWidth,
        HtmlRenderBoxStyle parentStyle, int depth, IElement? continuationTarget,
        int continuationLogicalCharacters, PagedFloatBoundary boundary,
        IReadOnlyList<HtmlFloatExclusion>? inheritedFloats, ICollection<HtmlFloatExclusion>? emittedFloats,
        InlineFloatContext? surroundingFloats, double gap, ConstrainedEntryRelayoutState state,
        out HtmlRenderFlowBlock reflowed) {
        reflowed = null!;
        ConstrainedSourceProjection context = _constrainedSourceProjection!;
        CheckCancellation();
        ChargeLayoutOperation("constrained fresh-entry relayout");
        // The same owner can retry only once on a recursion path: its content
        // then begins on a fresh page and the shared planner must make progress.
        if (!context.EntryRelayoutOwners.Add(element)) {
            context.FailedReason = "constrained-entry-without-fresh-page-progress";
            return false;
        }
        try {
            RestoreConstrainedEntryState(state, context);
            reflowed = LayoutElement(element, containingWidth, state.Style, parentStyle, depth,
                continuationTarget, continuationLogicalCharacters, boundary.Shift(gap), inheritedFloats,
                emittedFloats, surroundingFloats).AdjustLeadingFlowSpace(-gap);
            reflowed.PagedSourceEntryGap += gap;
            context.Changed = true;
            context.NeedsProjection = true;
            return true;
        } finally {
            context.EntryRelayoutOwners.Remove(element);
        }
    }

    private void RestoreConstrainedEntryState(ConstrainedEntryRelayoutState state,
        ConstrainedSourceProjection context) {
        CheckCancellation();
        ChargeLayoutOperations(1L + state.Placements.Length + state.Styles.Length
            + state.FloatOverhangs.Length + state.SemanticNodes.Length,
            "constrained entry state restoration");
        Restore(_normalFlowPlacements, state.Placements);
        Restore(_layoutStyles, state.Styles);
        Restore(_inlineFloatOverhangs, state.FloatOverhangs);
        Restore(_semanticNodeIds, state.SemanticNodes);
        _nextSemanticNodeId = state.SemanticOrder;
        _nextLogicalTextOrder = state.LogicalOrder;
        context.Changed = state.Changed;
        context.NeedsProjection = state.NeedsProjection;

        void Restore<T>(Dictionary<IElement, T> destination, KeyValuePair<IElement, T>[] entries) {
            destination.Clear();
            foreach (KeyValuePair<IElement, T> entry in entries) {
                CheckCancellation();
                destination.Add(entry.Key, entry.Value);
            }
        }
    }

    private sealed record ConstrainedEntryRelayoutState(HtmlRenderBoxStyle Style,
        KeyValuePair<IElement, NormalFlowPlacement>[] Placements,
        KeyValuePair<IElement, HtmlRenderBoxStyle>[] Styles,
        KeyValuePair<IElement, double>[] FloatOverhangs,
        KeyValuePair<IElement, int>[] SemanticNodes,
        int SemanticOrder, int LogicalOrder, bool Changed, bool NeedsProjection);
}
