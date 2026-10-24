using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Paints a positive-area solid background through the same box geometry for DOM and generated boxes.</summary>
    private void AddSolidBoxBackground(ICollection<HtmlRenderVisual> visuals, HtmlRenderBoxStyle style,
        double x, double y, double width, double height, HtmlRenderBorderInsets borderInsets,
        HtmlResolvedBorderRadii radii, string visualSourceDescription) {
        if (width <= 0.0001D || height <= 0.0001D) return;
        if (style.BackgroundColor.HasValue && style.BackgroundColor.Value.A > 0) {
            BackgroundPaintBox colorBox = ResolveBackgroundBox(
                style.BackgroundColorClip, x, y, width, height, borderInsets, style, radii);
            if (colorBox.Width > 0D && colorBox.Height > 0D) {
                OfficeShape fill = CreateBoxShape(colorBox.Width, colorBox.Height, colorBox.Radii);
                fill.FillColor = style.BackgroundColor;
                fill.StrokeWidth = 0D;
                bool canExtend = CanExtendBoxBackgroundThroughFragmentSlack(style, fill);
                visuals.Add(new HtmlRenderShape(fill, colorBox.X, colorBox.Y, visuals.Count, source: visualSourceDescription,
                    canExtendFragmentBackground: canExtend,
                    fragmentSideBorders: canExtend && (style.Display == "table" && style.BorderCollapse == "separate"
                        || CanExtendOrdinarySliceSides(style))
                        ? style.Borders : null));
            }
        }
    }
}
