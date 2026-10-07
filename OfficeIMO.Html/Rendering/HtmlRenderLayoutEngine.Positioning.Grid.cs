using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void RecordGridPositionedContainingRects(
        IElement container,
        HtmlRenderBoxStyle containerStyle,
        double contentWidth,
        double contentHeight,
        GridAxisLayout columns,
        GridAxisLayout rows,
        IReadOnlyDictionary<string, int> columnLineNames,
        IReadOnlyDictionary<string, int> rowLineNames,
        int explicitColumnCount,
        int explicitRowCount,
        int leadingColumnCount,
        int leadingRowCount) {
        if (!_localPositionedElements.TryGetValue(container, out List<PositionedElementRequest>? requests)) return;
        foreach (PositionedElementRequest request in requests.Where(item => ReferenceEquals(item.DirectParent, container))) {
            string source = HtmlRenderStyleResolver.DescribeSource(request.Element);
            GridAxisPlacement column = ParseGridAxisPlacement(
                request.Style.GridColumnStart,
                request.Style.GridColumnEnd,
                source,
                "grid-column",
                columnLineNames,
                explicitColumnCount,
                leadingColumnCount);
            GridAxisPlacement row = ParseGridAxisPlacement(
                request.Style.GridRowStart,
                request.Style.GridRowEnd,
                source,
                "grid-row",
                rowLineNames,
                explicitRowCount,
                leadingRowCount);
            ResolvePositionedGridAxis(columns, contentWidth, column.Start, column.Span, source, "grid-column", out double x, out double width);
            ResolvePositionedGridAxis(rows, contentHeight, row.Start, row.Span, source, "grid-row", out double y, out double height);
            _positionedContainingRects[request.Element] = new PositionedContainingRect(
                containerStyle.PaddingLeft + x,
                containerStyle.PaddingTop + y,
                width,
                height);
        }
    }

    private void ResolvePositionedGridAxis(
        GridAxisLayout axis,
        double fullSize,
        int? requestedStart,
        int requestedSpan,
        string source,
        string property,
        out double offset,
        out double size) {
        if (!requestedStart.HasValue) {
            offset = 0D;
            size = Math.Max(0.01D, fullSize);
            return;
        }

        int start = requestedStart.Value;
        int span = Math.Max(1, requestedSpan);
        if (start < 0 || start >= axis.Sizes.Count || start + span > axis.Sizes.Count) {
            ReportUnsupportedGridValue(source, property + " positioned area exceeded resolved tracks");
            start = Math.Max(0, Math.Min(axis.Sizes.Count - 1, start));
            span = Math.Max(1, Math.Min(span, axis.Sizes.Count - start));
        }
        offset = axis.Positions[start];
        size = axis.SpanSize(start, span);
    }
}
