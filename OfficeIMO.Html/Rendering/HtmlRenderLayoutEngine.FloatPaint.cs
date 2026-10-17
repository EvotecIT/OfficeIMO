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
                    if (fragment != null) result.Add(fragment.Translate(0D, 0D, result.Count));
                }
            }
            normal.Clear();
        }
        foreach (HtmlRenderVisual visual in visuals) {
            // Existing positioned layers retain their order relative to normal
            // phases. A descendant context also keeps its enclosing clip intact.
            if (HasNestedStackingContext(visual)) {
                FlushNormal();
                result.Add(visual.Translate(0D, 0D, result.Count));
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
        !visual.IsOutOfFlowPaint && (visual.PaintChildren is not { Count: > 0 } children
            || children.Any(ContainsInFlowPaint));

    private static bool ContainsDeferredFlowPaint(HtmlRenderVisual visual) =>
        visual.PaintPhase is HtmlRenderPaintPhase.Float or HtmlRenderPaintPhase.Outline
        || visual is HtmlRenderShape { IsCollapsedTableBorder: true }
        || visual.PaintPhase == HtmlRenderPaintPhase.Content && visual.StackingContext == null
            && GetFlowPaintChildren(visual)?.Any(ContainsDeferredFlowPaint) == true;

    private static IReadOnlyList<HtmlRenderVisual>? GetFlowPaintChildren(HtmlRenderVisual visual) {
        IReadOnlyList<HtmlRenderVisual>? children = visual is HtmlRenderEffectGroup ? null : GetGroupChildren(visual);
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
            && visual.PaintPhase == HtmlRenderPaintPhase.Content ? GetFlowPaintChildren(visual) : null;
        if (groupChildren != null) {
            HtmlRenderVisual[] children = groupChildren.Select(child => FilterFlowPaint(child, phase, projectionIdentities))
                .Where(child => child != null).Cast<HtmlRenderVisual>().ToArray();
            if (children.Length == 0) return null;
            HtmlRenderVisual fragment = visual is HtmlRenderLogicalTextGroup text
                ? visual.CopyStackingContextTo(text.ProjectPaint(children, 0D, 0D, visual.PaintOrder,
                    ownsLogicalText: phase == HtmlRenderPaintPhase.Content))
                : CloneGroupWithChildren(visual, children);
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
                HtmlRenderVisual root = GetFlowPaintChildren(visual) is { } children
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
