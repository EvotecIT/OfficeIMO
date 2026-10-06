using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<HtmlRenderVisual> SliceBlockVisuals(HtmlRenderFlowBlock block, double start, double end) {
        return SliceVisuals(block.Visuals, start, end);
    }

    private IReadOnlyList<HtmlRenderVisual> SliceVisuals(IEnumerable<HtmlRenderVisual> sourceVisuals, double start, double end) {
        var fragment = new List<HtmlRenderVisual>();
        foreach (HtmlRenderVisual visual in sourceVisuals) {
            int firstFragment = fragment.Count;
            SliceVisual(visual, start, end, fragment);
            for (int index = firstFragment; index < fragment.Count; index++) fragment[index].PaintPhase = visual.PaintPhase;
        }
        return fragment;
    }

    private void SliceVisual(HtmlRenderVisual visual, double start, double end, List<HtmlRenderVisual> fragment) {
        double visualTop = visual.LayoutY;
        double visualBottom = visual.LayoutY + visual.Height;
        double intersectionTop = Math.Max(start, visualTop);
        double intersectionBottom = Math.Min(end, visualBottom);
        if (intersectionBottom <= intersectionTop + 0.0001D) return;

        bool fullyContained = visualTop >= start - 0.0001D && visualBottom <= end + 0.0001D;
        if (fullyContained) {
            fragment.Add(visual.Translate(0D, -start, fragment.Count));
            return;
        }

        if (visual is HtmlRenderClipGroup clipGroup) {
            IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(clipGroup.Visuals, start, end);
            if (children.Count > 0) {
                fragment.Add(new HtmlRenderClipGroup(
                    clipGroup.ClipX,
                    clipGroup.ClipY - start,
                    clipGroup.ClipWidth,
                    clipGroup.ClipHeight,
                    clipGroup.ClipHorizontal,
                    clipGroup.ClipVertical,
                    children,
                    fragment.Count,
                    clipGroup.Source,
                    Math.Max(start, clipGroup.LayoutY) - start));
            }
            return;
        }

        if (visual is HtmlRenderSemanticGroup semanticGroup) {
            IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(semanticGroup.Visuals, start, end);
            if (children.Count > 0) {
                fragment.Add(new HtmlRenderSemanticGroup(
                    semanticGroup.Role,
                    semanticGroup.X,
                    semanticGroup.Y - start,
                    semanticGroup.Width,
                    Math.Max(0.01D, intersectionBottom - intersectionTop),
                    children,
                    fragment.Count,
                    semanticGroup.Source,
                    semanticGroup.ColumnSpan,
                    semanticGroup.RowSpan,
                    semanticGroup.HeaderScope,
                    semanticGroup.LayoutY - start,
                    semanticGroup.StructureElementKey));
            }
            return;
        }

        if (visual is HtmlRenderLayoutRegion layoutRegion) {
            IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(layoutRegion.Visuals, start, end);
            fragment.Add(new HtmlRenderLayoutRegion(
                layoutRegion.SourceKey,
                layoutRegion.RegionKind,
                layoutRegion.SourceText,
                layoutRegion.Position,
                layoutRegion.FloatSide,
                layoutRegion.ZIndex,
                layoutRegion.BackgroundLayerCount,
                layoutRegion.BoxShadowLayerCount,
                layoutRegion.BackgroundColor,
                layoutRegion.X,
                layoutRegion.Y - start,
                layoutRegion.Width,
                Math.Max(0.01D, intersectionBottom - intersectionTop),
                children,
                fragment.Count,
                layoutRegion.Source,
                layoutRegion.LayoutY - start));
            return;
        }

        if (visual is HtmlRenderBookmarkAnchor bookmarkAnchor) {
            if (bookmarkAnchor.LayoutY >= start - 0.0001D && bookmarkAnchor.LayoutY < end - 0.0001D) {
                fragment.Add(new HtmlRenderBookmarkAnchor(
                    bookmarkAnchor.SemanticNodeId,
                    bookmarkAnchor.Text,
                    bookmarkAnchor.X,
                    bookmarkAnchor.Y - start,
                    bookmarkAnchor.Width,
                    bookmarkAnchor.Height,
                    fragment.Count,
                    bookmarkAnchor.Source,
                    bookmarkAnchor.LayoutY - start));
            }
            return;
        }

        if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
            IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(logicalTextGroup.Visuals, start, end);
            if (children.Count > 0) {
                fragment.Add(new HtmlRenderLogicalTextGroup(
                    ResolveLogicalText(children, logicalTextGroup.Text),
                    logicalTextGroup.X,
                    logicalTextGroup.Y - start,
                    logicalTextGroup.Width,
                    Math.Max(0.01D, intersectionBottom - intersectionTop),
                    children,
                    fragment.Count,
                    logicalTextGroup.Source,
                    logicalTextGroup.LayoutY - start));
            }
            return;
        }

        if (visual is HtmlRenderEffectGroup effectGroup) {
            IReadOnlyList<HtmlRenderVisual> children = SliceVisuals(effectGroup.Visuals, start, end);
            if (children.Count > 0) {
                double translatedY = -start;
                OfficeTransform transform = OfficeTransform.Translate(0D, -translatedY)
                    .Then(effectGroup.Transform)
                    .Then(OfficeTransform.Translate(0D, translatedY));
                fragment.Add(new HtmlRenderEffectGroup(
                    effectGroup.X,
                    effectGroup.Y - start,
                    effectGroup.Width,
                    Math.Max(0.01D, intersectionBottom - intersectionTop),
                    transform,
                    effectGroup.Opacity,
                    children,
                    fragment.Count,
                    effectGroup.Source,
                    Math.Max(start, effectGroup.LayoutY) - start));
            }
            return;
        }

        if (visual is HtmlRenderImage
            || visual is HtmlRenderDrawing
            || visual is HtmlRenderImagePattern
            || visual is HtmlRenderPathClipGroup
            || visual is HtmlRenderShape) {
            fragment.Add(CreateVerticallyClippedVisualFragment(visual, start, intersectionTop, intersectionBottom, fragment.Count));
            return;
        }

        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.VisualFragmentUnsupported, "A visual crossing a forced page boundary could not be represented safely in the current fragment.", HtmlDiagnosticSeverity.Warning, visual.Source, visual.Kind.ToString());
    }

    private static HtmlRenderClipGroup CreateVerticallyClippedVisualFragment(
        HtmlRenderVisual visual,
        double fragmentStart,
        double intersectionTop,
        double intersectionBottom,
        int paintOrder) {
        double clipY = intersectionTop - fragmentStart;
        return new HtmlRenderClipGroup(
            visual.X,
            clipY,
            visual.Width,
            Math.Max(0.01D, intersectionBottom - intersectionTop),
            clipHorizontal: false,
            clipVertical: true,
            new[] { visual.Translate(0D, -fragmentStart, 0) },
            paintOrder,
            visual.Source,
            clipY);
    }

}
