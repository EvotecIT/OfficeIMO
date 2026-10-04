namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static HtmlRenderAvoidBreakRange? ResolveTrailingBoxKeepRange(HtmlRenderBoxStyle style,
        double contentHeight, double outerHeight, double contentY, IEnumerable<double> breakOffsets,
        IEnumerable<HtmlInlineBreakProgress> progress) {
        if (style.ExplicitHeight.HasValue && !style.AutoHeightFlexStretch || contentHeight <= 0.0001D
            || style.PaddingBottom + style.BorderBottomWidth <= 0.0001D) return null;

        // A final child's margin is spacing, not a content fragment. Keep the
        // preceding line with the container's bottom decoration across a page break.
        double finalContentEnd = progress.Where(item => item.IsBlockExit
                && item.PageStartDiscardableMargin > 0.0001D
                && Math.Abs(item.Offset + item.PageStartDiscardableMargin - contentHeight) <= 0.0001D)
            .Select(item => item.Offset).DefaultIfEmpty(contentHeight).Min();
        double finalContentStart = breakOffsets.Where(offset => offset < finalContentEnd - 0.0001D)
            .DefaultIfEmpty(0D).Max();
        double trailingBoxEnd = outerHeight - style.MarginBottom - contentY;
        return trailingBoxEnd > finalContentStart + 0.0001D
            ? new HtmlRenderAvoidBreakRange(finalContentStart, trailingBoxEnd)
            : null;
    }
}
