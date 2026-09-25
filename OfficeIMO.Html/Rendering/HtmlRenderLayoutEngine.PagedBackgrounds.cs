using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void ExtendPagedBodyBackgroundThroughFragmentSlack(
        HtmlRenderFlowBlock block,
        HtmlCssPageGeometry geometry,
        ICollection<HtmlRenderVisual> visuals,
        double fragmentBottom) {
        IElement? body = _document.Body;
        if (body == null || !ReferenceEquals(block.OwnerElement, body) || ReferenceEquals(_surfaceRootElement, body)) return;
        if (!_layoutStyles.TryGetValue(body, out HtmlRenderBoxStyle? style)
            || !style.PaintVisible
            || !style.BackgroundColor.HasValue
            || style.BackgroundColor.Value.A != byte.MaxValue
            || style.HasDeclaredBackgroundImage
            || style.Opacity < 1D
            || style.HorizontalInsets > 0.0001D
            || style.VerticalInsets > 0.0001D
            || style.HasBorderLayout
            || !string.Equals(style.BackgroundColorClip, "border-box", StringComparison.Ordinal)
            || !string.Equals(style.BorderRadius, "0", StringComparison.Ordinal)
            || style.BorderTopLeftRadius.Length > 0 || style.BorderTopRightRadius.Length > 0
            || style.BorderBottomLeftRadius.Length > 0 || style.BorderBottomRightRadius.Length > 0
            || !string.Equals(style.Transform, "none", StringComparison.Ordinal)
            || !string.Equals(style.IndividualScale, "none", StringComparison.Ordinal)
            || !string.Equals(style.ClipPath, "none", StringComparison.Ordinal)
            || Math.Abs(style.MarginLeft) > 0.0001D || Math.Abs(style.MarginRight) > 0.0001D) return;
        if (geometry.PrintProduction != null
            || geometry.Margins.Left > 0.0001D || geometry.Margins.Right > 0.0001D
            || geometry.Margins.Top > 0.0001D || geometry.Margins.Bottom > 0.0001D
            || Math.Abs(block.Width - geometry.ContentWidth) > 0.0001D) return;

        double slack = geometry.Height - fragmentBottom;
        if (slack <= 0.0001D) return;
        // A body fragment continues on the next page even when a safe break leaves unused
        // space here. Its background covers that slack, but not the last page after body ends.
        OfficeShape fill = OfficeShape.Rectangle(geometry.Width, slack);
        fill.FillColor = style.BackgroundColor.Value;
        fill.StrokeWidth = 0D;
        visuals.Add(new HtmlRenderShape(fill, 0D, fragmentBottom, int.MinValue + 1024,
            source: "body:fragment-background-continuation"));
    }
}
