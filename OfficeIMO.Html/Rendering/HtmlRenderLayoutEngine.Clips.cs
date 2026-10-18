using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static void AddBoxClipVisuals(
        ICollection<HtmlRenderVisual> target,
        IReadOnlyList<HtmlRenderVisual> children,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii,
        string source) {
        if (children.Count == 0 || width <= 0D || height <= 0D) return;
        HtmlResolvedBorderRadii normalized = radii.Normalize(width, height);
        if (normalized.IsZero && children.All(child => BoxClipContainsVisual(child, x, y, width, height))) {
            foreach (HtmlRenderVisual child in children) target.Add(child.Translate(0D, 0D, target.Count));
            return;
        }
        target.Add(new HtmlRenderPathClipGroup(
            x,
            y,
            CreateBoxClipPath(width, height, normalized),
            children,
            target.Count,
            source));
    }

    // Paint-neutral frames do not bound their descendants. Retain a real clip
    // for transformed or unknown visuals instead of treating their layout frame
    // as proof that all paint fits. Known leaves use their positioned envelope.
    private static bool BoxClipContainsVisual(HtmlRenderVisual visual, double x, double y, double width, double height) {
        if (visual is HtmlRenderSemanticGroup semantic) return semantic.Visuals.All(child => BoxClipContainsVisual(child, x, y, width, height));
        if (visual is HtmlRenderLogicalTextGroup logical) return logical.Visuals.All(child => BoxClipContainsVisual(child, x, y, width, height));
        if (visual is not (HtmlRenderShape or HtmlRenderText or HtmlRenderImage or HtmlRenderDrawing or HtmlRenderImagePattern or HtmlRenderClipGroup or HtmlRenderPathClipGroup)) return false;
        double stroke = visual is HtmlRenderShape shape ? shape.InnerShape.StrokeWidth / 2D : 0D;
        double topOverflow = visual is HtmlRenderText text ? text.PaintTopOverflow : stroke;
        return visual.X - stroke >= x && visual.Y - topOverflow >= y
            && visual.X + visual.Width + stroke <= x + width
            && visual.Y + visual.Height + stroke <= y + height;
    }

    private static OfficeClipPath CreateBoxClipPath(double width, double height, HtmlResolvedBorderRadii radii) =>
        radii.IsZero ? OfficeClipPath.Rectangle(width, height)
            : radii.IsUniformCircular
            ? OfficeClipPath.RoundedRectangle(width, height, radii.BoundedUniformRadius(width, height))
            : OfficeClipPath.Path(radii.CreatePathCommands(width, height));
}
