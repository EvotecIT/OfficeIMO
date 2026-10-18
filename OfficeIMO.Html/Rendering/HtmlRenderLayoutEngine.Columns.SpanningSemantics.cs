using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool IsFirstColumnSourceFragment(IElement element) =>
        !TryGetColumnSourceFragment(element, out ColumnSourceFragment fragment) || fragment.FirstFragment;

    private void RememberColumnSourceStyle(IElement element, HtmlRenderBoxStyle style) {
        if (IsFirstColumnSourceFragment(element) || !_layoutStyles.ContainsKey(element))
            _layoutStyles[element] = style.Clone();
    }

    private string ColumnSourceStructureKey(IElement element, HtmlRenderBoxStyle style) =>
        (style.Display == "contents" ? "html-display-contents:"
            : style.Display == "inline" ? "html-inline:" : "html-element:")
        + GetSemanticNodeId(element).ToString(System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>The extracted box retains its original logical ancestors and relative paint offsets.</summary>
    private HtmlRenderFlowBlock ApplyColumnSpannerAncestorSemantics(HtmlRenderFlowBlock block,
        IReadOnlyList<ColumnSpannerAncestor> ancestors) {
        for (int index = ancestors.Count - 1; index >= 0; index--) {
            CheckCancellation();
            ChargeLayoutOperation("column spanner semantic ancestor");
            ColumnSpannerAncestor ancestor = ancestors[index];
            HtmlRenderBoxStyle style = ancestor.Style;
            HtmlRenderSemanticGroupRole role = style.SemanticArtifact ? HtmlRenderSemanticGroupRole.Artifact
                : style.SemanticGroupRoleOverride ?? (TryResolveSemanticGroupRole(ancestor.Element.TagName,
                    out HtmlRenderSemanticGroupRole natural) ? natural : HtmlRenderSemanticGroupRole.Division);
            block = WrapSemanticBlock(block, role, HtmlRenderStyleResolver.DescribeSource(ancestor.Element),
                ColumnSourceStructureKey(ancestor.Element, style));
            block = ApplyPositioning(block, style, ancestor.ContainingWidth,
                ResolveContainingBlockHeight(ancestor.ParentStyle), HtmlRenderStyleResolver.DescribeSource(ancestor.Element));
        }
        return block;
    }

    /// <summary>Split ancestor pieces are composited together after their column projection.</summary>
    private HtmlRenderBoxStyle ColumnFragmentPaintStyle(IElement element, HtmlRenderBoxStyle style) {
        if ((!TryGetColumnSourceFragment(element, out _)
            && _columnSourceFlowContext?.OpacityOwners.ContainsKey(element) != true)
            || !HasColumnSourceOpacity(style)) return style;
        HtmlRenderBoxStyle paint = style.Clone();
        paint.Opacity = 1D;
        paint.OpacityWasSpecified = false;
        return paint;
    }

    private static bool HasColumnSourceOpacity(HtmlRenderBoxStyle style) =>
        style.Display != "contents" && style.OpacityWasSpecified
        && style.UnsupportedOpacity.Length == 0 && style.Opacity < 1D;

    /// <summary>Defers one finite overflow owner's opacity until its paint has been routed through columns.</summary>
    private void RegisterColumnOverflowOpacity(IElement element, HtmlRenderBoxStyle style, int depth) {
        if (_columnSourceFlowContext == null || !HasColumnSourceOpacity(style)) return;
        CheckCancellation();
        ChargeLayoutOperation("column overflow opacity owner");
        _columnSourceFlowContext.OpacityOwners[element] = new ColumnSourceOpacityOwner(element, style.Clone(), depth);
    }

    private sealed class ColumnSourceOpacityOwner {
        internal ColumnSourceOpacityOwner(IElement element, HtmlRenderBoxStyle style, int depth) {
            Element = element;
            Style = style;
            Depth = depth;
        }
        internal IElement Element { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal int Depth { get; }
    }

    private HtmlRenderFlowBlock ApplyColumnSourceOpacity(HtmlRenderFlowBlock content,
        IReadOnlyList<MultiColumnSpanPartition>? partitions = null) {
        var ancestors = new Dictionary<IElement, ColumnSourceOpacityOwner>();
        if (_columnSourceFlowContext != null) {
            foreach (ColumnSourceOpacityOwner owner in _columnSourceFlowContext.OpacityOwners.Values) {
                CheckCancellation();
                ChargeLayoutOperation("column overflow opacity discovery");
                ancestors[owner.Element] = owner;
            }
        }
        foreach (MultiColumnSpanPartition partition in partitions ?? Array.Empty<MultiColumnSpanPartition>()) {
            foreach (ColumnSpannerAncestor ancestor in partition.SpannerAncestors) {
                CheckCancellation();
                ChargeLayoutOperation("column ancestor opacity discovery");
                if (HasColumnSourceOpacity(ancestor.Style)) ancestors[ancestor.Element] =
                    new ColumnSourceOpacityOwner(ancestor.Element, ancestor.Style, ancestor.Depth);
            }
        }
        IReadOnlyList<HtmlRenderVisual> visuals = content.Visuals;
        foreach (ColumnSourceOpacityOwner ancestor in ancestors.Values.OrderByDescending(item => item.Depth)) {
            string key = ColumnSourceStructureKey(ancestor.Element, ancestor.Style);
            var inside = new List<HtmlRenderVisual>();
            var outside = new List<HtmlRenderVisual>();
            int insertion = -1;
            foreach (HtmlRenderVisual visual in visuals) {
                Split(visual, key, out HtmlRenderVisual? selected, out HtmlRenderVisual? remainder);
                if (selected != null) {
                    if (insertion < 0) insertion = outside.Count;
                    inside.Add(selected);
                }
                if (remainder != null) outside.Add(remainder);
            }
            if (inside.Count == 0) continue;
            IReadOnlyList<HtmlRenderVisual> painted = ReplaceDescendantFormFieldsForPaintEffect(inside,
                "ancestor-opacity=" + ancestor.Style.Opacity.ToString(System.Globalization.CultureInfo.InvariantCulture));
            outside.Insert(insertion, new HtmlRenderEffectGroup(0D, 0D, content.Width, content.PagedPaintExtent,
                OfficeTransform.Identity, ancestor.Style.Opacity, painted, insertion,
                HtmlRenderStyleResolver.DescribeSource(ancestor.Element)).WithStackingContext(
                    new HtmlRenderStackingContext(ancestor.Style.Position is "relative" or "sticky"
                        ? ResolvePositionedZIndex(ancestor.Element, ancestor.Style) : 0,
                        GetPositionedSourceOrder(ancestor.Element))));
            visuals = outside.Select((visual, index) => visual.Translate(0D, 0D, index)).ToArray();
        }
        return content.WithVisuals(visuals);

        void Split(HtmlRenderVisual visual, string key,
            out HtmlRenderVisual? selected, out HtmlRenderVisual? remainder) {
            CheckCancellation();
            ChargeLayoutOperation("column ancestor opacity projection");
            if (visual is HtmlRenderSemanticGroup semantic && semantic.StructureElementKey == key) {
                selected = visual;
                remainder = null;
                return;
            }
            if (visual.PaintChildren is not { } children) {
                selected = null;
                remainder = visual;
                return;
            }
            var included = new List<HtmlRenderVisual>();
            var retained = new List<HtmlRenderVisual>();
            foreach (HtmlRenderVisual child in children) {
                Split(child, key, out HtmlRenderVisual? childSelected, out HtmlRenderVisual? childRemainder);
                if (childSelected != null) included.Add(childSelected);
                if (childRemainder != null) retained.Add(childRemainder);
            }
            selected = included.Count == 0 ? null
                : visual.ProjectPaintChildren(included, 0D, 0D, visual.PaintOrder);
            remainder = retained.Count == 0 ? null
                : included.Count == 0 ? visual : visual.ProjectPaintChildren(retained, 0D, 0D, visual.PaintOrder);
        }
    }
}
