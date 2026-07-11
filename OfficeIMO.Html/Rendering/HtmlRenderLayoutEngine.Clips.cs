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
        if (children.Count == 0) return;
        HtmlResolvedBorderRadii normalized = radii.Normalize(width, height);
        if (normalized.IsZero) {
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

    private static OfficeClipPath CreateBoxClipPath(double width, double height, HtmlResolvedBorderRadii radii) =>
        radii.IsUniformCircular
            ? OfficeClipPath.RoundedRectangle(width, height, radii.UniformRadius)
            : OfficeClipPath.Path(radii.CreatePathCommands(width, height));
}
