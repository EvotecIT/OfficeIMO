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
        IReadOnlyDictionary<string, GridAreaDefinition> areas,
        IReadOnlyDictionary<string, int> columnLineNames,
        IReadOnlyDictionary<string, int> rowLineNames) {
        if (!_localPositionedElements.TryGetValue(container, out List<PositionedElementRequest>? requests)) return;
        foreach (PositionedElementRequest request in requests.Where(item => ReferenceEquals(item.DirectParent, container))) {
            string source = HtmlRenderStyleResolver.DescribeSource(request.Element);
            int? requestedRow = null;
            int? requestedColumn = null;
            int rowSpan = 1;
            int columnSpan = 1;
            string areaName = request.Style.GridArea;
            if (areaName != "auto" && areaName.IndexOf('/') < 0 && areas.TryGetValue(areaName, out GridAreaDefinition? area)) {
                requestedRow = area.Row;
                requestedColumn = area.Column;
                rowSpan = area.RowSpan;
                columnSpan = area.ColumnSpan;
            } else {
                GridAxisPlacement column = ParseGridAxisPlacement(
                    request.Style.GridColumnStart,
                    request.Style.GridColumnEnd,
                    source,
                    "grid-column",
                    columnLineNames);
                GridAxisPlacement row = ParseGridAxisPlacement(
                    request.Style.GridRowStart,
                    request.Style.GridRowEnd,
                    source,
                    "grid-row",
                    rowLineNames);
                requestedColumn = column.Start;
                requestedRow = row.Start;
                columnSpan = column.Span;
                rowSpan = row.Span;
            }

            ResolvePositionedGridAxis(columns, contentWidth, requestedColumn, columnSpan, source, "grid-column", out double x, out double width);
            ResolvePositionedGridAxis(rows, contentHeight, requestedRow, rowSpan, source, "grid-row", out double y, out double height);
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
