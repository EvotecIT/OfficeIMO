namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Retains constrained in-flow source without changing the used CSS box or sibling advance.</summary>
    private double ResolveConstrainedBlockPaintExtent(HtmlRenderBoxStyle style, double extent,
        double contentY, double contentHeight, HtmlInlineLayout? inline,
        IReadOnlyList<FlowPaintLayer> children) {
        if (_options.Mode != HtmlRenderMode.Paged || style.OverflowY != "visible"
            || !HasConstrainedBlockHeight(style)) return extent;

        double contentExtent = Math.Max(contentHeight, inline?.PagedPaintExtent ?? 0D);
        foreach (FlowPaintLayer child in children) {
            CheckCancellation();
            ChargeLayoutOperation("constrained block source extent");
            contentExtent = Math.Max(contentExtent, child.Y + ResolveNormalFlowPaintExtent(child.Block));
        }
        return Math.Max(extent, contentY + contentExtent);
    }

    /// <summary>Exposes only the known trailing allocation space of an authored height constraint.</summary>
    private void AppendConstrainedBlockAllocationTail(HtmlRenderBoxStyle style, double boxHeight,
        double contentY, double contentHeight, IReadOnlyList<FlowPaintLayer> children,
        ICollection<HtmlRenderBreakRange> ranges, double fullPageHeight) {
        if (_options.Mode != HtmlRenderMode.Paged || !HasConstrainedBlockHeight(style)) return;
        double boxEnd = style.MarginTop + boxHeight;
        double contentEnd = contentY + Math.Max(0D, contentHeight);
        if (boxEnd <= contentEnd + 0.0001D) return;
        CheckCancellation();
        ChargeLayoutOperation("constrained block allocation interval");
        // A negative-margin child can still paint into this allocation tail.
        // Qualify its table/keep/source ownership through the existing shared
        // interval owner before the enclosing paint filter checks all carriers.
        IReadOnlyList<HtmlRenderBreakRange> safe = CollectSafeParallelBreakRanges(
            children.Select(child => (child.Block, child.Y)), boxEnd - contentY,
            new[] { new HtmlRenderBreakRange(contentEnd - contentY, boxEnd - contentY) }, fullPageHeight);
        foreach (HtmlRenderBreakRange range in safe) ranges.Add(range);
    }

    private static bool HasConstrainedBlockHeight(HtmlRenderBoxStyle style) =>
        style.ExplicitHeight.HasValue || style.MinHeight.HasValue
        || style.MaxHeight.HasValue || style.AspectRatio.HasValue;

    private static HtmlRenderAvoidBreakRange? ResolveTrailingBoxKeepRange(HtmlRenderBoxStyle style,
        double contentHeight, double outerHeight, double contentY, IEnumerable<double> breakOffsets,
        IEnumerable<HtmlInlineBreakProgress> progress) {
        if (style.ExplicitHeight.HasValue && !style.AutoHeightLayoutStretch || contentHeight <= 0.0001D
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
