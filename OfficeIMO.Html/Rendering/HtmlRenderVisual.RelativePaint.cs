namespace OfficeIMO.Html;

public abstract partial class HtmlRenderVisual {
    // Only CSS relative positioning sets this value. Glyph placement, shadows and
    // transforms also move paint, so Y - LayoutY cannot identify relative paint.
    internal double RelativePaintOffsetY { get; private set; }
    internal bool IsOutOfFlowPaint { get; private set; }

    internal HtmlRenderVisual TranslateRelativePaint(double offsetX, double offsetY, int paintOrder) {
        HtmlRenderVisual result = TranslatePaint(offsetX, offsetY, paintOrder);
        Mark(result);
        return result;

        void Mark(HtmlRenderVisual visual) {
            if (visual.IsOutOfFlowPaint) return;
            visual.RelativePaintOffsetY += offsetY;
            if (visual.PaintChildren is { } children) {
                foreach (HtmlRenderVisual child in children) Mark(child);
            }
        }
    }

    // Called only on a freshly translated positioned layer.
    internal HtmlRenderVisual IdentifyOutOfFlowPaint() {
        Mark(this);
        return this;

        static void Mark(HtmlRenderVisual visual) {
            visual.IsOutOfFlowPaint = true;
            if (visual.PaintChildren is { } children) {
                foreach (HtmlRenderVisual child in children) Mark(child);
            }
        }
    }

    internal IReadOnlyList<HtmlRenderVisual>? PaintChildren => this switch {
        HtmlRenderClipGroup clip => clip.Visuals,
        HtmlRenderPathClipGroup path => path.Visuals,
        HtmlRenderEffectGroup effect => effect.Visuals,
        HtmlRenderSemanticGroup semantic => semantic.Visuals,
        HtmlRenderLayoutRegion region => region.Visuals,
        HtmlRenderLogicalTextGroup logical => logical.Visuals,
        HtmlRenderFormField formField => formField.Visuals,
        _ => null
    };

    internal HtmlRenderVisual ProjectPaintChildren(IEnumerable<HtmlRenderVisual> children,
        double offsetX, double offsetY, int paintOrder, bool ownsLogicalText = true) =>
        CopyStackingContextTo(this switch {
            HtmlRenderClipGroup clip => clip.ProjectPaint(children, offsetX, offsetY, paintOrder,
                clip.IsFlowFragment ? RelativePaintOffsetY : 0D),
            HtmlRenderPathClipGroup path => path.ProjectPaint(children, offsetX, offsetY, paintOrder),
            HtmlRenderEffectGroup effect => effect.ProjectPaint(children, offsetX, offsetY, paintOrder),
            HtmlRenderSemanticGroup semantic => semantic.ProjectPaint(children, offsetX, offsetY, paintOrder),
            HtmlRenderLayoutRegion region => region.ProjectPaint(children, offsetX, offsetY, paintOrder),
            HtmlRenderLogicalTextGroup logical => logical.ProjectPaint(children, offsetX, offsetY, paintOrder, ownsLogicalText),
            HtmlRenderFormField formField => formField.ProjectPaint(children, offsetX, offsetY, paintOrder),
            _ => throw new InvalidOperationException("A leaf visual cannot project child paint.")
        });
}
