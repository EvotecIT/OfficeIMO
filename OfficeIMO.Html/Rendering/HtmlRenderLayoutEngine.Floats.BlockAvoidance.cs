namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static HtmlRenderBoxStyle AvoidActiveFloatsForFormattingContext(
        HtmlRenderBoxStyle style,
        double containingWidth,
        double flowHeight,
        IReadOnlyList<HtmlFloatExclusion>? activeFloats,
        double? measuredHeight = null) {
        if (activeFloats == null || activeFloats.Count == 0
            || !EstablishesFloatContainingBlock(style)) return style;

        double originalTop = flowHeight + style.MarginTop;
        double top = originalTop;
        double requiredWidth = style.ExplicitWidth ?? style.MinWidth ?? 0.01D;
        if (style.MaxWidth.HasValue) requiredWidth = Math.Min(requiredWidth, style.MaxWidth.Value);
        if (style.MinWidth.HasValue) requiredWidth = Math.Max(requiredWidth, style.MinWidth.Value);
        if (!style.BorderBox) requiredWidth += style.HorizontalInsets;
        requiredWidth += style.MarginLeft + style.MarginRight;
        double extent = measuredHeight ?? (style.ExplicitHeight.HasValue
                ? Math.Max(0.01D, Math.Max(style.ExplicitHeight.Value, style.MinHeight ?? 0D) + style.VerticalInsets)
                : Math.Max(0.01D, (style.MinHeight ?? 0D) + style.VerticalInsets));
        for (int attempt = 0; attempt <= activeFloats.Count; attempt++) {
            double left = 0D;
            double right = containingWidth;
            double nextTop = double.PositiveInfinity;
            bool intersectsFloat = false;
            foreach (HtmlFloatExclusion exclusion in activeFloats) {
                if (exclusion.Y >= top + extent - 0.0001D || exclusion.Bottom <= top + 0.0001D) continue;
                intersectsFloat = true;
                if (exclusion.Side == "left") left = Math.Max(left, exclusion.Right);
                else if (exclusion.Side == "right") right = Math.Min(right, exclusion.X);
                nextTop = Math.Min(nextTop, exclusion.Bottom);
            }

            // An over-wide box still belongs below the float once no float
            // intersects it; it may then overflow the containing width.
            if (!intersectsFloat || right - left >= requiredWidth - 0.0001D && right - left > 1D) {
                if (top == originalTop && left <= 0D && right >= containingWidth) return style;
                HtmlRenderBoxStyle adjusted = style.Clone();
                adjusted.MarginTop += top - originalTop;
                adjusted.MarginLeft += left;
                adjusted.MarginRight += containingWidth - right;
                return adjusted;
            }

            if (double.IsPositiveInfinity(nextTop) || nextTop <= top + 0.0001D) break;
            top = nextTop;
        }
        return style;
    }
}
