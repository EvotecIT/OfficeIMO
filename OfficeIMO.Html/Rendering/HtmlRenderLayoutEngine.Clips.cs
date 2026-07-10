using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static void AddRoundedClipVisuals(
        ICollection<HtmlRenderVisual> target,
        IReadOnlyList<HtmlRenderVisual> children,
        double x,
        double y,
        double width,
        double height,
        double radius,
        string source) {
        if (children.Count == 0) return;
        double boundedRadius = Math.Min(Math.Max(0D, radius), Math.Min(width, height) / 2D);
        if (boundedRadius <= 0.0001D) {
            foreach (HtmlRenderVisual child in children) target.Add(child.Translate(0D, 0D, target.Count));
            return;
        }
        target.Add(new HtmlRenderPathClipGroup(
            x,
            y,
            OfficeClipPath.RoundedRectangle(width, height, boundedRadius),
            children,
            target.Count,
            source));
    }
}
