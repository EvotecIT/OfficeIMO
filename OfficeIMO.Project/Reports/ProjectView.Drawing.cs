using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Project;

/// <summary>A rendered report page in point coordinates, with its source row and bucket ranges.</summary>
public sealed class ProjectViewPage {
    internal ProjectViewPage(OfficeDrawing drawing, int rowOffset, int rowCount, int bucketOffset, int bucketCount) {
        Drawing = drawing; RowOffset = rowOffset; RowCount = rowCount; BucketOffset = bucketOffset; BucketCount = bucketCount;
    }
    /// <summary>Independent editable drawing owned by this rendering call.</summary>
    public OfficeDrawing Drawing { get; }
    /// <summary>First source row shown.</summary>
    public int RowOffset { get; }
    /// <summary>Number of source rows shown.</summary>
    public int RowCount { get; }
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
        if (Kind == ProjectViewKind.Network) return RenderNetwork(cancellationToken);
        bool table = Kind == ProjectViewKind.Table;
        double width = Layout.PageWidth - 2 * Layout.Margin;
        int rowsPerPage = Math.Max(1, (int)((Layout.PageHeight - 2 * Layout.Margin - 115) / 30));
        double labelWidth = table ? width : Kind == ProjectViewKind.Timeline ? 0 : Math.Min(width * .55, Math.Max(220, Columns.Count * 80));
        int bucketsPerPage = table ? Math.Max(1, Buckets.Count) : Math.Max(1, (int)((width - labelWidth) / 66));
        int vertical = Math.Max(1, (Rows.Count + rowsPerPage - 1) / rowsPerPage);
        int horizontal = table ? 1 : Math.Max(1, (Buckets.Count + bucketsPerPage - 1) / bucketsPerPage);
        if ((long)vertical * horizontal > Layout.MaxPages) throw new InvalidOperationException("Report exceeds MaxPages.");
        var pages = new List<ProjectViewPage>();
        for (int v = 0; v < vertical; v++) for (int h = 0; h < horizontal; h++) {
            cancellationToken.ThrowIfCancellationRequested();
            int rowStart = v * rowsPerPage, rowCount = Math.Min(rowsPerPage, Rows.Count - rowStart);
            int bucketStart = h * bucketsPerPage, bucketCount = table ? 0 : Math.Min(bucketsPerPage, Buckets.Count - bucketStart);
            var drawing = NewPage(pages.Count + 1, vertical * horizontal);
            double x = Layout.Margin, y = Layout.Margin + 58;
            if (table) DrawTable(drawing, rowStart, rowCount, x, y, width);
            else DrawTimeGrid(drawing, rowStart, rowCount, bucketStart, bucketCount, x, y, width, labelWidth);
            if (Rows.Count == 0) Text(drawing, "No rows match this selection.", x, y + 35, width, 25, 12, Muted);
            pages.Add(new ProjectViewPage(drawing, rowStart, rowCount, bucketStart, bucketCount));
        }
        return pages.AsReadOnly();
    }

    private OfficeDrawing NewPage(int number, int count) {
        var drawing = new OfficeDrawing(Layout.PageWidth, Layout.PageHeight);
        Rect(drawing, 0, 0, Layout.PageWidth, Layout.PageHeight, OfficeColor.White);
        Text(drawing, Title, Layout.Margin, Layout.Margin, Layout.PageWidth - 2 * Layout.Margin, 26, 20, Ink, true);
        Text(drawing, Kind + (Buckets.Count == 0 ? "" : "  ·  " + Buckets[0].Start.ToString("dd MMM yyyy", CultureInfo.InvariantCulture)
            + " – " + Buckets[Buckets.Count - 1].Finish.ToString("dd MMM yyyy", CultureInfo.InvariantCulture)), Layout.Margin, Layout.Margin + 30,
            Layout.PageWidth - 2 * Layout.Margin, 18, 10, Muted);
        string legend = Kind == ProjectViewKind.TaskUsage || Kind == ProjectViewKind.ResourceUsage || Kind == ProjectViewKind.ResourceHistogram
            ? "Work hours · includes actual and overtime · excludes material quantities"
            : Kind == ProjectViewKind.Network ? "Arrows: dependencies · page references: linked task location · red: critical · dark: summary"
            : Kind == ProjectViewKind.Gantt ? "Blue: scheduled · red: critical · dark: summary · gray: baseline · arrows: links within page (all links in Network)"
            : "Blue: scheduled · red: critical · dark: summary · gray: baseline";
        if (Layout.ShowLegend) Text(drawing, legend, Layout.Margin, Layout.PageHeight - Layout.Margin - 28, Layout.PageWidth - 2 * Layout.Margin, 16, 9, Muted);
        Text(drawing, "Page " + number + " / " + count, Layout.Margin, Layout.PageHeight - Layout.Margin - 12, Layout.PageWidth - 2 * Layout.Margin, 12, 8, Muted);
        return drawing;
    }

    private void DrawTable(OfficeDrawing drawing, int offset, int count, double x, double y, double width) {
        double cell = width / Columns.Count;
        Rect(drawing, x, y, width, 26, Ink);
        for (int c = 0; c < Columns.Count; c++) Text(drawing, ColumnTitle(Columns[c]), x + c * cell + 4, y + 5, cell - 8, 18, 9, OfficeColor.White, true);
        for (int r = 0; r < count; r++) {
            var row = Rows[offset + r]; double top = y + 26 + r * 30;
            if (r % 2 == 0) Rect(drawing, x, top, width, 30, OfficeColor.ParseHex("#F4F7FA"));
            for (int c = 0; c < Columns.Count; c++) Text(drawing, row.GetText(Columns[c]), x + c * cell + 4, top + 6, cell - 8, 22, 9, row.IsCritical ? Critical : Ink, row.IsSummary);
        }
    }

    private void DrawTimeGrid(OfficeDrawing drawing, int rowOffset, int rowCount, int bucketOffset, int bucketCount,
        double x, double y, double width, double labelWidth) {
        bool timeline = Kind == ProjectViewKind.Timeline;
        Rect(drawing, x, y, width, 26, Ink);
        var weights = Columns.Select(c => c == ProjectViewColumn.Name ? 2.5 : c == ProjectViewColumn.Uid ? .7 : 1.5).ToArray();
        double totalWeight = weights.Sum(), left = x;
        for (int c = 0; !timeline && c < Columns.Count; c++) {
            double w = labelWidth * weights[c] / totalWeight;
            Text(drawing, ColumnTitle(Columns[c]), left + 4, y + 5, w - 8, 18, 9, OfficeColor.White, true); left += w;
        }
        double cell = (width - labelWidth) / Math.Max(1, bucketCount), chartX = x + labelWidth;
        for (int b = 0; b < bucketCount; b++) Text(drawing, Buckets[bucketOffset + b].Start.ToString("dd MMM yy", CultureInfo.InvariantCulture), chartX + b * cell + 3, y + 6, cell - 6, 17, 9, OfficeColor.White);
        decimal maximum = Kind == ProjectViewKind.ResourceHistogram ? Rows.SelectMany(r => r.BucketWorkHours).DefaultIfEmpty(0).Max() : 0;
        for (int r = 0; r < rowCount; r++) {
            var row = Rows[rowOffset + r]; double top = y + 26 + r * 30;
            if (r % 2 == 0) Rect(drawing, x, top, width, 30, OfficeColor.ParseHex("#F4F7FA"));
            left = x;
            for (int c = 0; !timeline && c < Columns.Count; c++) {
                double w = labelWidth * weights[c] / totalWeight;
                string value = Columns[c] == ProjectViewColumn.Start ? row.Start?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? ""
                    : Columns[c] == ProjectViewColumn.Finish ? row.Finish?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "" : row.GetText(Columns[c]);
                Text(drawing, value, left + 4, top + 3, w - 8, 17, 9, row.IsCritical ? Critical : Ink, row.IsSummary); left += w;
            }
            if (timeline) Text(drawing, string.Join(" · ", Columns.Select(row.GetText)), x + 4, top + 1, width - 8, 15, 9, row.IsSummary ? Ink : row.IsCritical ? Critical : Ink, row.IsSummary);
            if (!timeline && row.Group.Length > 0) Text(drawing, row.Group, x + 5, top + 18, labelWidth - 10, 11, 7, Muted);
            for (int b = 0; b < bucketCount; b++) Rect(drawing, chartX + b * cell, top, .5, 30, Grid);
            if (bucketCount == 0) continue;
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
                DrawBar(drawing, row.BaselineStart, row.BaselineFinish, bucketOffset, bucketCount, chartX, top + (timeline ? 28 : 24), cell, timeline ? 2 : 3, Muted);
                DrawBar(drawing, row.Start, row.Finish, bucketOffset, bucketCount, chartX, top + (timeline ? 18 : 7), cell, timeline ? 8 : row.IsSummary ? 8 : 13, row.IsSummary ? Ink : row.IsCritical ? Critical : Accent);
            }
        }
        if (Kind == ProjectViewKind.Gantt && bucketCount > 0) DrawGanttLinks(drawing, rowOffset, rowCount, bucketOffset, bucketCount, chartX, y + 26, cell);
    }

    private void DrawBar(OfficeDrawing drawing, DateTime? start, DateTime? finish, int offset, int count,
        double x, double y, double cellWidth, double height, OfficeColor color) {
        var first = Buckets[offset].Start; var last = Buckets[offset + count - 1].Finish;
        if (!start.HasValue || !finish.HasValue || finish < first || start >= last || finish < start) return;
        double width = count * cellWidth, left = TimePosition(start.Value, offset, count, cellWidth), right = TimePosition(finish.Value, offset, count, cellWidth);
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
    private static void Rect(OfficeDrawing drawing, double x, double y, double width, double height, OfficeColor fill) {
        var shape = OfficeShape.Rectangle(width, height); shape.FillColor = fill; shape.StrokeColor = null;
        drawing.AddShape(shape, x, y);
    }
    private static void Text(OfficeDrawing drawing, string text, double x, double y, double width, double height, double size, OfficeColor color, bool bold = false) =>
        drawing.AddText(text, x, y, width, height, new OfficeFontInfo("Arial", size, bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular), color);
}
