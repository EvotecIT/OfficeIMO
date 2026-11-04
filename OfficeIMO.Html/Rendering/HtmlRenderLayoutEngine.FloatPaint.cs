namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Block backgrounds paint below floats and in-flow content. Outlines paint
    // after normal-flow siblings, within the same existing stacking boundaries.
    private static IEnumerable<HtmlRenderVisual> OrderFloatPaint(IReadOnlyList<HtmlRenderVisual> visuals) {
        if (!visuals.Any(ContainsDeferredFlowPaint)) return visuals;
        var result = new List<HtmlRenderVisual>();
        var normal = new List<HtmlRenderVisual>();
        var projectionIdentities = new Dictionary<HtmlRenderVisual, object>();
        void FlushNormal() {
            foreach (HtmlRenderPaintPhase phase in new[] { HtmlRenderPaintPhase.BlockBackground,
                HtmlRenderPaintPhase.Float, HtmlRenderPaintPhase.Content, HtmlRenderPaintPhase.Outline }) {
                foreach (HtmlRenderVisual visual in normal) {
                    HtmlRenderVisual? fragment = FilterFlowPaint(visual, phase, projectionIdentities);
                    if (fragment != null) result.Add(fragment.WithPaintOrder(result.Count, owned: !ReferenceEquals(fragment, visual)));
                }
            }
            normal.Clear();
        }
        foreach (HtmlRenderVisual visual in visuals) {
            // Existing positioned layers retain their order relative to normal
            // phases. A descendant context also keeps its enclosing clip intact.
            if (HasNestedStackingContext(visual) || ContainsPositionedFlowPaint(visual)) {
                FlushNormal();
                result.Add(visual.WithPaintOrder(result.Count, owned: false));
            } else {
                normal.Add(visual);
            }
        }
        FlushNormal();
        return result;
    }

    // Semantic, clip and effect envelopes can contain only positioned paint.
    // Wrapping that paint does not make it a normal-flow content carrier.
    private static bool ContainsInFlowPaint(HtmlRenderVisual visual) =>
        !visual.IsOutOfFlowPaint && (visual.InvariantPaintChildren is not { Count: > 0 } children
            || children.Any(ContainsInFlowPaint));

    // Local positioned layers already belong to their host's ordered bands.
    // Transparent envelopes retain those bands; independent boxes keep their
    // positioned descendants inside their own atomic or float paint phase.
    private static bool ContainsPositionedFlowPaint(HtmlRenderVisual visual) =>
        visual.IsOutOfFlowPaint
        || visual.PaintPhase == HtmlRenderPaintPhase.Content
            && GetFlowPaintChildren(visual, invariantOnly: true)?.Any(ContainsPositionedFlowPaint) == true;

    private static bool ContainsDeferredFlowPaint(HtmlRenderVisual visual) =>
        visual.PaintPhase is HtmlRenderPaintPhase.BlockBackground or HtmlRenderPaintPhase.Float or HtmlRenderPaintPhase.Outline
        || visual is HtmlRenderShape { IsCollapsedTableBorder: true }
        || visual.PaintPhase == HtmlRenderPaintPhase.Content && visual.StackingContext == null
            && GetFlowPaintChildren(visual, invariantOnly: true)?.Any(ContainsDeferredFlowPaint) == true;

    private static IReadOnlyList<HtmlRenderVisual>? GetFlowPaintChildren(HtmlRenderVisual visual, bool invariantOnly = false) {
        IReadOnlyList<HtmlRenderVisual>? children = visual is HtmlRenderEffectGroup ? null
            : invariantOnly && visual is HtmlRenderSemanticGroup semantic ? semantic.InvariantVisuals
            : GetGroupChildren(visual);
        // Empty groups can own logical text or paint-neutral pagination anchors.
        // Keep them as one content leaf rather than filtering away every phase.
        return children is { Count: > 0 } ? children : null;
    }

    private static HtmlRenderVisual? FilterFlowPaint(
        HtmlRenderVisual visual, HtmlRenderPaintPhase phase,
        IDictionary<HtmlRenderVisual, object> projectionIdentities) {
        // Effects, positioned contexts, floats and independent formatting boxes
        // own their descendants' order and remain atomic in the parent phases.
        IReadOnlyList<HtmlRenderVisual>? groupChildren = visual.StackingContext == null
            && visual.PaintPhase == HtmlRenderPaintPhase.Content ? GetFlowPaintChildren(visual, invariantOnly: true) : null;
        if (groupChildren != null) {
            // Phases do not change under translation, so a pending move is filtered in source
            // coordinates and carried by the fragment instead of copying every child here.
            HtmlRenderSemanticGroup? pendingGroup = null;
            HtmlRenderOffset[]? pendingOffsets = null;
            if (visual is HtmlRenderSemanticGroup semantic) {
                if (semantic.TryGetPendingTranslation(out IReadOnlyList<HtmlRenderVisual> source, out HtmlRenderOffset[] offsets)) {
                    pendingGroup = semantic;
                    pendingOffsets = offsets;
                    groupChildren = source;
                } else {
                    groupChildren = semantic.Visuals;
                }
            }
            var filtered = new List<HtmlRenderVisual>(groupChildren.Count);
            for (int index = 0; index < groupChildren.Count; index++) {
                HtmlRenderVisual? child = FilterFlowPaint(groupChildren[index], phase, projectionIdentities);
                if (child == null) continue;
                // Number children as the caller's renumbering copy would, so the fragment needs none.
                // Pending children are numbered by position when the fragment is read.
                filtered.Add(pendingOffsets != null ? child
                    : child.WithPaintOrder(filtered.Count, owned: !ReferenceEquals(child, groupChildren[index])));
            }
            if (filtered.Count == 0) return null;
            // Group constructors copy their children, so the working list can be passed directly.
            HtmlRenderVisual fragment = pendingGroup != null
                ? visual.CopyStackingContextTo(pendingGroup.WithPendingChildren(filtered, pendingOffsets!))
                : visual is HtmlRenderLogicalTextGroup text
                ? visual.CopyStackingContextTo(text.ProjectPaint(filtered, 0D, 0D, visual.PaintOrder,
                    ownsLogicalText: phase == HtmlRenderPaintPhase.Content))
                : CloneGroupWithChildren(visual, filtered);
            if (!projectionIdentities.TryGetValue(visual, out object? identity)) {
                identity = visual.PaintProjectionIdentity ?? new object();
                projectionIdentities.Add(visual, identity);
            }
            fragment.IdentifyPaintProjection(identity);
            fragment.PaintPhase = phase;
            return fragment;
        }
        HtmlRenderPaintPhase effective = visual.PaintPhase == HtmlRenderPaintPhase.Atomic
            ? HtmlRenderPaintPhase.Content : visual.PaintPhase;
        return effective == phase ? visual : null;
    }

    /// <summary>Orders a box before an atomic formatting or effect envelope hides its phases.</summary>
    private static HtmlRenderFlowBlock OrderBoxFlowPaint(HtmlRenderFlowBlock block) {
        IEnumerable<HtmlRenderVisual> ordered = OrderFloatPaint(block.Visuals);
        bool changed = !ReferenceEquals(ordered, block.Visuals);
        if (block.ContinuationGroups.Count == 0 && block.TrailingGroups.Count == 0) {
            return changed ? block.WithVisuals(ordered) : block;
        }
        HtmlRenderContinuationGroup[] continuations = block.ContinuationGroups.Select(group => {
            IEnumerable<HtmlRenderVisual> groupPaint = OrderFloatPaint(group.Visuals);
            if (ReferenceEquals(groupPaint, group.Visuals)) return group;
            changed = true;
            return new HtmlRenderContinuationGroup(group.StartsAfter, group.EndsAt, group.Height, groupPaint);
        }).ToArray();
        HtmlRenderTrailingGroup[] trailing = block.TrailingGroups.Select(group => {
            IEnumerable<HtmlRenderVisual> groupPaint = OrderFloatPaint(group.Visuals);
            if (ReferenceEquals(groupPaint, group.Visuals)) return group;
            changed = true;
            return new HtmlRenderTrailingGroup(group.StartsAt, group.ContentEndsAt,
                group.SourceEndsAt, group.Height, groupPaint);
        }).ToArray();
        return changed ? block.WithVisuals(ordered, continuationGroups: continuations, trailingGroups: trailing) : block;
    }

    /// <summary>Closes the same formatting boundary on ordinary and detached repeated fragments.</summary>
    private static void CloseBoxFlowPaint(HtmlRenderFlowBlock block) {
        if (block.ContinuationGroups.Count == 0 && block.TrailingGroups.Count == 0) {
            foreach (HtmlRenderVisual visual in block.Visuals) visual.CloseFlowPaint();
            return;
        }
        // Visual classes retain reference equality. A shared root crosses this
        // boundary once even if more than one fragment channel retains it.
        var closed = new HashSet<HtmlRenderVisual>();
        foreach (HtmlRenderVisual visual in block.Visuals
            .Concat(block.ContinuationGroups.SelectMany(group => group.Visuals))
            .Concat(block.TrailingGroups.SelectMany(group => group.Visuals))) {
            if (closed.Add(visual)) visual.CloseFlowPaint();
        }
    }

    /// <summary>Recombines the paint phases of one source block's body and repeated table groups.</summary>
    private static IEnumerable<HtmlRenderVisual> OrderRepeatedTableFragmentPaint(IReadOnlyList<HtmlRenderVisual> visuals) {
        var roots = new List<HtmlRenderVisual>(visuals.Count);
        foreach (HtmlRenderVisual visual in visuals) {
            // A block's root envelope is atomic relative to outside siblings.
            // Here all roots belong to that same block; its nested independent
            // formatting/effect/stacking contexts retain their existing order.
            if (visual.PaintPhase == HtmlRenderPaintPhase.Atomic && visual.StackingContext == null
                && visual.AtomicSourcePaintPhase is { } sourcePhase) {
                HtmlRenderVisual root = visual is HtmlRenderSemanticGroup semantic
                    && semantic.TryGetPendingTranslation(out IReadOnlyList<HtmlRenderVisual> source, out HtmlRenderOffset[] offsets)
                    && source.Count > 0
                    ? visual.CopyStackingContextTo(semantic.WithPendingChildren(source, offsets))
                    : GetFlowPaintChildren(visual) is { } children
                    ? CloneGroupWithChildren(visual, children)
                    : visual.Translate(0D, 0D, visual.PaintOrder);
                root.PaintPhase = sourcePhase;
                roots.Add(root);
            } else {
                roots.Add(visual);
            }
        }
        return OrderFloatPaint(roots);
    }

    private static List<HtmlRenderVisual> OrderPageFloatPaint(List<HtmlRenderVisual> visuals) {
        if (!visuals.Any(ContainsDeferredFlowPaint)) return visuals;
        var normal = visuals.Where(visual => visual.PaintOrder >= 0 && visual.PaintOrder < 1000000000).OrderBy(visual => visual.PaintOrder).ToList();
        return visuals.Where(visual => visual.PaintOrder < 0).OrderBy(visual => visual.PaintOrder)
            .Concat(OrderFloatPaint(normal))
            .Concat(visuals.Where(visual => visual.PaintOrder >= 1000000000).OrderBy(visual => visual.PaintOrder)).ToList();
    }
}
