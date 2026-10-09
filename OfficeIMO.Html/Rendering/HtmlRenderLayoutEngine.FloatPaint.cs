namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Non-positioned block backgrounds paint below floats; in-flow text paints
    // above them. Semantic fragments retain their structure identity when split.
    private static IEnumerable<HtmlRenderVisual> OrderFloatPaint(IReadOnlyList<HtmlRenderVisual> visuals) {
        if (!visuals.Any(ContainsFloatPaint)) return visuals;
        var result = new List<HtmlRenderVisual>();
        foreach (HtmlRenderPaintPhase phase in new[] { HtmlRenderPaintPhase.BlockBackground, HtmlRenderPaintPhase.Float, HtmlRenderPaintPhase.Content }) {
            foreach (HtmlRenderVisual visual in visuals) {
                HtmlRenderVisual? fragment = FilterFloatPaint(visual, phase);
                if (fragment != null) result.Add(fragment.Translate(0D, 0D, result.Count));
            }
        }
        return result;
    }

    private static bool ContainsFloatPaint(HtmlRenderVisual visual) =>
        visual.PaintPhase == HtmlRenderPaintPhase.Float
        || visual.PaintPhase == HtmlRenderPaintPhase.Content && visual is HtmlRenderSemanticGroup group
            && group.StackingContext == null && group.Visuals.Any(ContainsFloatPaint);

    private static HtmlRenderVisual? FilterFloatPaint(HtmlRenderVisual visual, HtmlRenderPaintPhase phase) {
        // A positioned/effect stacking context already owns its descendants'
        // order. Splitting it into root float phases can cover its own text.
        if (visual.StackingContext == null
            && visual.PaintPhase == HtmlRenderPaintPhase.Content && visual is HtmlRenderSemanticGroup group) {
            HtmlRenderVisual[] children = group.Visuals.Select(child => FilterFloatPaint(child, phase))
                .Where(child => child != null).Cast<HtmlRenderVisual>().ToArray();
            if (children.Length == 0) return null;
            HtmlRenderVisual fragment = group.CopyStackingContextTo(
                group.ProjectPaint(children, 0D, 0D, group.PaintOrder));
            fragment.PaintPhase = phase;
            return fragment;
        }
        HtmlRenderPaintPhase effective = visual.PaintPhase == HtmlRenderPaintPhase.Atomic ? HtmlRenderPaintPhase.Content : visual.PaintPhase;
        return effective == phase ? visual : null;
    }

    private static List<HtmlRenderVisual> OrderPageFloatPaint(List<HtmlRenderVisual> visuals) {
        if (!visuals.Any(ContainsFloatPaint)) return visuals;
        var normal = visuals.Where(visual => visual.PaintOrder >= 0 && visual.PaintOrder < 1000000000).OrderBy(visual => visual.PaintOrder).ToList();
        return visuals.Where(visual => visual.PaintOrder < 0).OrderBy(visual => visual.PaintOrder)
            .Concat(OrderFloatPaint(normal))
            .Concat(visuals.Where(visual => visual.PaintOrder >= 1000000000).OrderBy(visual => visual.PaintOrder)).ToList();
    }
}
