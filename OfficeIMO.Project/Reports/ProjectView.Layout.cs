using OfficeIMO.Drawing;

namespace OfficeIMO.Project;

public sealed partial class ProjectView {
    private string LegendText => Kind == ProjectViewKind.TaskUsage || Kind == ProjectViewKind.ResourceUsage || Kind == ProjectViewKind.ResourceHistogram
        ? "Work hours · includes actual and overtime · excludes material quantities"
        : Kind == ProjectViewKind.Network ? "Arrows: dependencies · p: linked task page · red: critical · dark: summary"
        : Kind == ProjectViewKind.Gantt ? "Blue: scheduled · red: critical · dark: summary/progress · gray: baseline · dashed: status date · arrows: links within page (all links in Network)"
        : "Blue: scheduled · red: critical · dark: summary · gray: baseline";

    private double MeasureTextHeight(string text, double size, double width, bool bold = false) {
        var layout = OfficeTextLayoutEngine.LayoutTextBlock(text, size, width, Layout.PageHeight, 1.3, size,
            (value, fontSize) => MeasureWidth(value, fontSize, bold), true);
        if (layout.Clipped) throw new InvalidOperationException("Report text cannot fit on a page. Increase the page width or height.");
        return layout.Height;
    }

    private double TitleHeight => Math.Max(26, MeasureTextHeight(Title, 20, Layout.PageWidth - 2 * Layout.Margin, true));
    private double TitleAreaHeight => TitleHeight + 32;
    private double FooterHeight => 18 + (Layout.ShowLegend ? MeasureTextHeight(LegendText, 9, Layout.PageWidth - 2 * Layout.Margin) + 6 : 0);
    private double BodyHeight => Layout.PageHeight - 2 * Layout.Margin - TitleAreaHeight - FooterHeight - 16;

    private double[] ColumnWidths(double width) {
        var minimums = Columns.Select(c => c == ProjectViewColumn.Uid ? 38d : c == ProjectViewColumn.Name ? 110d : 82d).ToArray();
        double minimum = minimums.Sum();
        var weights = Columns.Select(c => c == ProjectViewColumn.Name ? 4d : 1d).ToArray();
        double total = weights.Sum();
        return minimums.Select((value, index) => width < minimum ? value * width / minimum : value + (width - minimum) * weights[index] / total).ToArray();
    }

    private string DisplayText(ProjectViewRow row, ProjectViewColumn column) => Kind != ProjectViewKind.Table && (column == ProjectViewColumn.Start || column == ProjectViewColumn.Finish)
        ? (column == ProjectViewColumn.Start ? row.Start : row.Finish)?.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture) ?? ""
        : row.GetText(column);

    private double MeasureHeaderHeight(double width) {
        if (Kind == ProjectViewKind.Timeline) return 26;
        var widths = ColumnWidths(width);
        double height = 26;
        for (int c = 0; c < Columns.Count; c++) {
            var text = OfficeTextLayoutEngine.LayoutTextBlock(ColumnTitle(Columns[c]), 10, Math.Max(1, widths[c] - 12), Layout.PageHeight - 2 * Layout.Margin - 160,
                1.3, 10, (value, size) => MeasureWidth(value, size, true), true);
            if (text.Clipped) throw new InvalidOperationException("Report column headers cannot fit on a page. Increase the page or column width.");
            height = Math.Max(height, text.Height + 10);
        }
        return height;
    }

    private double[] MeasureRows(double labelWidth, double availableHeight, CancellationToken token) {
        var widths = ColumnWidths(labelWidth);
        var heights = new double[Rows.Count];
        for (int r = 0; r < Rows.Count; r++) {
            token.ThrowIfCancellationRequested();
            var row = Rows[r]; double height = 36;
            bool timeline = Kind == ProjectViewKind.Timeline;
            for (int c = 0; c < (timeline ? 1 : Columns.Count); c++) {
                string value = timeline ? string.Join(" · ", Columns.Select(column => DisplayText(row, column))) : DisplayText(row, Columns[c]);
                var text = OfficeTextLayoutEngine.LayoutTextBlock(value, 11, Math.Max(1, (timeline ? labelWidth : widths[c]) - 12), availableHeight - (timeline ? 40 : 20),
                    1.3, 11, (textValue, size) => MeasureWidth(textValue, size, row.IsSummary), true);
                if (text.Clipped) throw new InvalidOperationException("A report label cannot fit on a page. Increase the page or column width.");
                height = Math.Max(height, text.Height + (timeline ? 36 : 16) + (!timeline && Kind != ProjectViewKind.Table && row.Group.Length > 0 ? 14 : 0));
            }
            heights[r] = height;
        }
        return heights;
    }
}
