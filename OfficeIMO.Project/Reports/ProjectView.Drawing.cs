using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Project;

/// <summary>A rendered report page in point coordinates, with its source row and bucket ranges.</summary>
public sealed class ProjectViewPage {
    internal ProjectViewPage(OfficeDrawing drawing, int rowOffset, int rowCount, int bucketOffset, int bucketCount, int[]? rowIndices = null) {
        Drawing = drawing; RowOffset = rowOffset; RowCount = rowCount; BucketOffset = bucketOffset; BucketCount = bucketCount;
        RowIndices = Array.AsReadOnly(rowIndices ?? Enumerable.Range(rowOffset, rowCount).ToArray());
    }
    /// <summary>Independent editable drawing owned by this rendering call.</summary>
    public OfficeDrawing Drawing { get; }
    /// <summary>Lowest source row index shown. Use RowIndices for the exact selection in dependency diagrams.</summary>
    public int RowOffset { get; }
    /// <summary>Number of source rows shown.</summary>
    public int RowCount { get; }
    /// <summary>Exact source row indices in visual order. Dependency diagrams can reorder and select noncontiguous rows.</summary>
    public IReadOnlyList<int> RowIndices { get; }
    /// <summary>First time bucket shown.</summary>
    public int BucketOffset { get; }
    /// <summary>Number of time buckets shown.</summary>
    public int BucketCount { get; }
}

public sealed partial class ProjectView {
    private static readonly OfficeColor Ink = OfficeColor.ParseHex("#183047");
    private static readonly OfficeColor Accent = OfficeColor.ParseHex("#247C9C");
    private static readonly OfficeColor Critical = OfficeColor.ParseHex("#B93848");
    private static readonly OfficeColor Muted = OfficeColor.ParseHex("#64748B");
    private static readonly OfficeColor Grid = OfficeColor.ParseHex("#E3EAF0");

    /// <summary>Renders bounded, horizontally and vertically paginated report drawings using the shared Core drawing model.</summary>
    public IReadOnlyList<ProjectViewPage> Render(CancellationToken cancellationToken = default) {
        if (_typography == null) return Render(OfficeRenderingProfile.Managed, cancellationToken);
        if (Kind == ProjectViewKind.Network) return RenderNetwork(cancellationToken);
        bool table = Kind == ProjectViewKind.Table;
        double width = Layout.PageWidth - 2 * Layout.Margin;
        double labelWidth = table ? width : Kind == ProjectViewKind.Timeline ? 0 : Math.Min(width * .55, Math.Max(220, Columns.Count * 80));
        double headerHeight = MeasureHeaderHeight(labelWidth);
        var heights = MeasureRows(Kind == ProjectViewKind.Timeline ? width : labelWidth, BodyHeight - headerHeight, cancellationToken);
        var rowPages = OfficeTablePagination.Paginate(heights, BodyHeight, headerHeight, Layout.MaxPages, cancellationToken);
        int bucketsPerPage = table ? Math.Max(1, Buckets.Count) : Math.Max(1, (int)((width - labelWidth) / 66));
        int vertical = rowPages.Count;
        int horizontal = table ? 1 : Math.Max(1, (Buckets.Count + bucketsPerPage - 1) / bucketsPerPage);
        if ((long)vertical * horizontal > Layout.MaxPages) throw new InvalidOperationException("Report exceeds MaxPages.");
        var pages = new List<ProjectViewPage>();
        for (int v = 0; v < vertical; v++) for (int h = 0; h < horizontal; h++) {
            cancellationToken.ThrowIfCancellationRequested();
            int rowStart = rowPages[v].RowOffset, rowCount = rowPages[v].RowCount;
            var pageHeights = heights.Skip(rowStart).Take(rowCount).ToArray();
            int bucketStart = h * bucketsPerPage, bucketCount = table ? 0 : Math.Min(bucketsPerPage, Buckets.Count - bucketStart);
            var drawing = NewPage(pages.Count + 1, vertical * horizontal, rowPages[v].Height + (Rows.Count == 0 ? 36 : 0));
            double x = Layout.Margin, y = Layout.Margin + TitleAreaHeight;
            if (table) DrawTable(drawing, rowStart, rowCount, x, y, width, headerHeight, pageHeights);
            else DrawTimeGrid(drawing, rowStart, rowCount, bucketStart, bucketCount, x, y, width, labelWidth, headerHeight, pageHeights);
            if (Rows.Count == 0) Text(drawing, "No rows match this selection.", x, y + headerHeight + 10, width, 25, 12, Muted);
            pages.Add(new ProjectViewPage(drawing, rowStart, rowCount, bucketStart, bucketCount));
        }
        return pages.AsReadOnly();
    }

    private OfficeDrawing NewPage(int number, int count, double? contentHeight = null) {
        double pageHeight = Layout.FitPageHeightToContent && contentHeight.HasValue
            ? Math.Min(Layout.PageHeight, contentHeight.Value + 2 * Layout.Margin + TitleAreaHeight + FooterHeight + 16) : Layout.PageHeight;
        var drawing = new OfficeDrawing(Layout.PageWidth, pageHeight);
        if (_typography != null) {
            drawing.Fonts.AddRange(_typography.Fonts);
            drawing.TextShapingProvider = _typography.TextShapingProvider;
            drawing.TextShapingLanguage = _typography.TextShapingLanguage;
        }
        Rect(drawing, 0, 0, Layout.PageWidth, pageHeight, OfficeColor.White);
        Text(drawing, Title, Layout.Margin, Layout.Margin, Layout.PageWidth - 2 * Layout.Margin, TitleHeight, 20, Ink, true, true);
        Text(drawing, KindTitle(Kind) + (Buckets.Count == 0 ? "" : "  ·  " + Buckets[0].Start.ToString("dd MMM yyyy", CultureInfo.InvariantCulture)
            + " – " + Buckets[Buckets.Count - 1].Finish.ToString("dd MMM yyyy", CultureInfo.InvariantCulture)), Layout.Margin, Layout.Margin + TitleHeight + 4,
            Layout.PageWidth - 2 * Layout.Margin, 18, 10, Muted);
        if (Layout.ShowLegend) Text(drawing, LegendText, Layout.Margin, pageHeight - Layout.Margin - FooterHeight, Layout.PageWidth - 2 * Layout.Margin, FooterHeight - 18, 9, Muted, wrap: true);
        Text(drawing, "Page " + number + " / " + count, Layout.Margin, pageHeight - Layout.Margin - 12, Layout.PageWidth - 2 * Layout.Margin, 12, 9, Muted);
        return drawing;
    }

    private void DrawTable(OfficeDrawing drawing, int offset, int count, double x, double y, double width, double headerHeight, double[] heights) {
        var widths = ColumnWidths(width); double left = x;
        Rect(drawing, x, y, width, headerHeight, Ink);
        for (int c = 0; c < Columns.Count; c++) { Text(drawing, ColumnTitle(Columns[c]), left + 6, y + 5, widths[c] - 12, headerHeight - 10, 10, OfficeColor.White, true, true); left += widths[c]; }
        double top = y + headerHeight;
        for (int r = 0; r < count; r++) {
            var row = Rows[offset + r]; left = x;
            if (r % 2 == 0) Rect(drawing, x, top, width, heights[r], OfficeColor.ParseHex("#F4F7FA"));
            for (int c = 0; c < Columns.Count; c++) { Text(drawing, row.GetText(Columns[c]), left + 6, top + 8, widths[c] - 12, heights[r] - 16, 11, Ink, row.IsSummary, true); left += widths[c]; }
            top += heights[r];
        }
    }

    private void DrawTimeGrid(OfficeDrawing drawing, int rowOffset, int rowCount, int bucketOffset, int bucketCount,
        double x, double y, double width, double labelWidth, double headerHeight, double[] heights) {
        bool timeline = Kind == ProjectViewKind.Timeline;
        Rect(drawing, x, y, width, headerHeight, Ink);
        var weights = ColumnWidths(labelWidth);
        double totalWeight = weights.Sum(), left = x;
        for (int c = 0; !timeline && c < Columns.Count; c++) {
            double w = labelWidth * weights[c] / totalWeight;
            Text(drawing, ColumnTitle(Columns[c]), left + 6, y + 5, w - 12, headerHeight - 10, 10, OfficeColor.White, true, true); left += w;
        }
        double cell = (width - labelWidth) / Math.Max(1, bucketCount), chartX = x + labelWidth;
        for (int b = 0; b < bucketCount; b++) Text(drawing, Buckets[bucketOffset + b].Start.ToString("dd MMM yy", CultureInfo.InvariantCulture), chartX + b * cell + 3, y + 6, cell - 6, 17, 9, OfficeColor.White);
        decimal maximum = Kind == ProjectViewKind.ResourceHistogram ? Rows.SelectMany(r => r.BucketWorkHours).DefaultIfEmpty(0).Max() : 0;
        double top = y + headerHeight;
        for (int r = 0; r < rowCount; r++) {
            var row = Rows[rowOffset + r];
            if (r % 2 == 0) Rect(drawing, x, top, width, heights[r], OfficeColor.ParseHex("#F4F7FA"));
            left = x;
            for (int c = 0; !timeline && c < Columns.Count; c++) {
                double w = labelWidth * weights[c] / totalWeight;
                string value = DisplayText(row, Columns[c]);
                Text(drawing, value, left + 6, top + 8, w - 12, heights[r] - 16 - (row.Group.Length > 0 ? 14 : 0), 11, Ink, row.IsSummary, true); left += w;
            }
            if (timeline) Text(drawing, string.Join(" · ", Columns.Select(column => DisplayText(row, column))), x + 6, top + 6, width - 12, heights[r] - 36, 11, Ink, row.IsSummary, true);
            if (!timeline && row.Group.Length > 0) Text(drawing, row.Group, x + 6, top + heights[r] - 16, labelWidth - 12, 14, 9, Muted);
            for (int b = 0; b < bucketCount; b++) Rect(drawing, chartX + b * cell, top, .5, heights[r], Grid);
            if (bucketCount == 0) { top += heights[r]; continue; }
            if (Kind == ProjectViewKind.TaskUsage || Kind == ProjectViewKind.ResourceUsage || Kind == ProjectViewKind.ResourceHistogram) {
                for (int b = 0; b < bucketCount; b++) {
                    decimal hours = row.BucketWorkHours[bucketOffset + b];
                    if (Kind == ProjectViewKind.ResourceHistogram && maximum > 0 && hours > 0) {
                        double barHeight = (double)(hours / maximum) * 24;
                        Rect(drawing, chartX + b * cell + 3, top + 28 - barHeight, cell - 6, barHeight, OfficeColor.ParseHex("#B9DCE8"));
                    }
                    Text(drawing, hours.ToString("0.##", CultureInfo.InvariantCulture), chartX + b * cell + 4, top + 8, cell - 8, 18, 10, Ink);
                }
            } else {
                double center = timeline ? heights[r] - 18 : heights[r] / 2 - 3;
                double barHeight = timeline || row.IsSummary ? 8 : 13;
                DrawBar(drawing, row.BaselineStart, row.BaselineFinish, bucketOffset, bucketCount, chartX, top + center + 10, cell, 3, Muted);
                DrawBar(drawing, row.Start, row.Finish, bucketOffset, bucketCount, chartX, top + center - barHeight / 2, cell, barHeight, row.IsSummary ? Ink : row.IsCritical ? Critical : Accent);
                if (row.IsSummary) DrawSummaryCaps(drawing, row, bucketOffset, bucketCount, chartX, top + center, cell);
                else if (Layout.ShowProgress && row.PercentComplete > 0 && row.Start.HasValue && row.Finish > row.Start) {
                    long elapsed = (long)((row.Finish.Value.Ticks - row.Start.Value.Ticks) * Math.Min(100, row.PercentComplete.Value) / 100m);
                    DrawBar(drawing, row.Start, row.Start.Value.AddTicks(elapsed), bucketOffset, bucketCount, chartX, top + center - 2, cell, 4, Ink);
                }
            }
            top += heights[r];
        }
        if (Kind == ProjectViewKind.Gantt && bucketCount > 0) DrawGanttLinks(drawing, rowOffset, rowCount, bucketOffset, bucketCount, chartX, y + headerHeight, cell, heights);
        if ((Kind == ProjectViewKind.Gantt || timeline) && bucketCount > 0 && Layout.StatusDate.HasValue
            && Layout.StatusDate >= Buckets[bucketOffset].Start && Layout.StatusDate < Buckets[bucketOffset + bucketCount - 1].Finish) {
            double markerX = chartX + TimePosition(Layout.StatusDate.Value, bucketOffset, bucketCount, cell);
            var marker = OfficeShape.Line(markerX, y + headerHeight, markerX, top);
            marker.StrokeColor = Accent; marker.StrokeWidth = 1; marker.StrokeDashStyle = OfficeStrokeDashStyle.Dash;
            drawing.AddShape(marker, markerX, y + headerHeight);
        }
    }

    private void DrawBar(OfficeDrawing drawing, DateTime? start, DateTime? finish, int offset, int count,
        double x, double y, double cellWidth, double height, OfficeColor color) {
        var first = Buckets[offset].Start; var last = Buckets[offset + count - 1].Finish;
        if (!start.HasValue || !finish.HasValue || finish < first || start >= last || finish < start) return;
        double width = count * cellWidth, left = TimePosition(start.Value, offset, count, cellWidth), right = TimePosition(finish.Value, offset, count, cellWidth);
        if (start == finish) {
            double size = Math.Max(8, height);
            double center = Math.Max(size / 2, Math.Min(width - size / 2, left));
            var diamond = OfficeShape.Polygon(new OfficePoint(size / 2, 0), new OfficePoint(size, size / 2),
                new OfficePoint(size / 2, size), new OfficePoint(0, size / 2));
            diamond.FillColor = color; diamond.StrokeColor = null;
            drawing.AddShape(diamond, x + center - size / 2, y + (height - size) / 2);
            return;
        }
        if (right <= 0 && finish != start) return;
        double barWidth = Math.Min(width - left, Math.Max(2, right - left));
        if (barWidth > 0) Rect(drawing, x + left, y, barWidth, height, color);
    }

    private double TimePosition(DateTime value, int offset, int count, double cellWidth) {
        if (value <= Buckets[offset].Start) return 0;
        if (value >= Buckets[offset + count - 1].Finish) return count * cellWidth;
        for (int i = 0; i < count; i++) {
            var bucket = Buckets[offset + i];
            if (value < bucket.Finish) return (i + (value.Ticks - bucket.Start.Ticks) / (double)(bucket.Finish.Ticks - bucket.Start.Ticks)) * cellWidth;
        }
        return count * cellWidth;
    }

    /// <summary>User-facing label for a typed column.</summary>
    public static string ColumnTitle(ProjectViewColumn column) => column switch {
        ProjectViewColumn.WorkHours => "Work (hours)", ProjectViewColumn.PercentComplete => "% complete",
        ProjectViewColumn.Uid => "UID", _ when Enum.IsDefined(typeof(ProjectViewColumn), column) => column.ToString(),
        _ => throw new ArgumentOutOfRangeException(nameof(column))
    };
    /// <summary>User-facing name of a report layout.</summary>
    public static string KindTitle(ProjectViewKind kind) => kind switch {
        ProjectViewKind.TaskUsage => "Task usage", ProjectViewKind.ResourceUsage => "Resource usage",
        ProjectViewKind.ResourceHistogram => "Resource histogram",
        _ when Enum.IsDefined(typeof(ProjectViewKind), kind) => kind.ToString(),
        _ => throw new ArgumentOutOfRangeException(nameof(kind))
    };
    private static void Rect(OfficeDrawing drawing, double x, double y, double width, double height, OfficeColor fill) {
        var shape = OfficeShape.Rectangle(width, height); shape.FillColor = fill; shape.StrokeColor = null;
        drawing.AddShape(shape, x, y);
    }
    private static void Text(OfficeDrawing drawing, string text, double x, double y, double width, double height, double size, OfficeColor color, bool bold = false, bool wrap = false) =>
        drawing.AddText(text, x, y, width, height, new OfficeFontInfo("Arial", size, bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular), color, lineHeight: size * 1.3, wrapText: wrap);
}
