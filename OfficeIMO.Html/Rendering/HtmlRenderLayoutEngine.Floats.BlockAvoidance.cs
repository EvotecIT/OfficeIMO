namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static HtmlRenderBoxStyle AvoidActiveFloatsForFormattingContext(
        HtmlRenderBoxStyle style,
        double containingWidth,
        double flowHeight,
        IReadOnlyList<HtmlFloatExclusion>? activeFloats) {
        if (activeFloats == null || activeFloats.Count == 0
            || !EstablishesFloatContainingBlock(style)
            || style.ExplicitWidth.HasValue) return style;

        double top = flowHeight + style.MarginTop;
        double left = 0D;
        double right = containingWidth;
        foreach (HtmlFloatExclusion exclusion in activeFloats) {
            if (exclusion.Y > top + 0.0001D || exclusion.Bottom <= top + 0.0001D) continue;
            if (exclusion.Side == "left") left = Math.Max(left, exclusion.Right);
            else if (exclusion.Side == "right") right = Math.Min(right, exclusion.X);
        }

        if (right - left <= 1D || left <= 0D && right >= containingWidth) return style;
        HtmlRenderBoxStyle adjusted = style.Clone();
        adjusted.MarginLeft += left;
        adjusted.MarginRight += containingWidth - right;
        return adjusted;
    }
}
