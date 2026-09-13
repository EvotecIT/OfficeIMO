using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.Project;

namespace OfficeIMO.Workflows;

public static partial class ProjectReportWorkflow {
    private static void AddTableSlides(PowerPointPresentation presentation, ProjectView view, TablePart part, OfficeRasterCanvas measurement, CancellationToken token) {
        double width = view.PageWidth - 2 * view.PageMargin;
        double availableHeight = view.PageHeight - 2 * view.PageMargin - 85;
        var weights = ColumnWeights(part.Headers);
        var widths = weights.Select(w => width * w / weights.Sum()).ToArray();
        double MeasureRow(string[] values, bool bold) {
            double height = 26;
            for (int c = 0; c < values.Length; c++) {
                token.ThrowIfCancellationRequested();
                var layout = OfficeTextLayoutEngine.LayoutTextBlock(values[c], 12, Math.Max(1, widths[c] - 18), availableHeight - 12,
                    1.5, 12, (text, size) => measurement.MeasureText(text, size, "Arial", bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular), true);
                if (layout.Clipped) throw new InvalidOperationException("An editable PowerPoint table cell cannot fit on a slide. Use a wider page or export the complete table to Word or Excel.");
                height = Math.Max(height, layout.Height + 12);
            }
            return height;
        }
        double headerHeight = MeasureRow(part.Headers, true);
        var heights = part.Rows.Select(row => MeasureRow(row, false)).ToArray();
        foreach (var page in OfficeTablePagination.Paginate(heights, availableHeight, headerHeight, view.MaxPages, token)) {
            token.ThrowIfCancellationRequested();
            if (presentation.Slides.Count >= view.MaxPages) throw new InvalidOperationException("Editable presentation exceeds MaxPages.");
            int start = page.RowOffset, count = page.RowCount; double height = page.Height;
            var slide = presentation.AddSlide();
            var title = slide.AddTextBoxPoints(view.Title, view.PageMargin, view.PageMargin, width, 32);
            title.FontName = "Arial"; title.FontSize = 22; title.Color = "183047";
            var subtitle = slide.AddTextBoxPoints(part.Title + (start > 0 ? " (continued)" : ""), view.PageMargin, view.PageMargin + 36, width, 25);
            subtitle.FontName = "Arial"; subtitle.FontSize = 13; subtitle.Color = "64748B";
            var table = slide.AddTable(count + 1, part.Headers.Length);
            table.LeftPoints = view.PageMargin; table.TopPoints = view.PageMargin + 72; table.WidthPoints = width; table.HeightPoints = height;
            table.SetColumnWidthsPoints(widths); table.HeaderRow = true;
            var rows = table.RowItems;
            for (int r = 0; r <= count; r++) {
                rows[r].HeightPoints = r == 0 ? headerHeight : heights[start + r - 1];
                for (int c = 0; c < part.Headers.Length; c++) {
                    var cell = rows[r].Cells[c]; cell.Text = r == 0 ? part.Headers[c] : part.Rows[start + r - 1][c];
                    cell.FontSize = 12; cell.FontName = "Arial"; cell.Bold = r == 0; cell.TextAutoFit = PowerPointTextAutoFit.None;
                    cell.FillColor = r == 0 ? "183047" : r % 2 == 1 ? "F4F7FA" : "FFFFFF";
                    cell.Color = r == 0 ? "FFFFFF" : "183047"; cell.BorderColor = "E3EAF0";
                }
            }
            slide.Notes.Text = "Editable report data. Source revision " + view.ModelRevision + ". " + view.Kind;
        }
    }
}
