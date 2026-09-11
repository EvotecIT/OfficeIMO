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
        int start = 0;
        do {
            token.ThrowIfCancellationRequested();
            if (presentation.Slides.Count >= view.MaxPages) throw new InvalidOperationException("Editable presentation exceeds MaxPages.");
            int count = 0; double height = headerHeight;
            while (start + count < heights.Length && height + heights[start + count] <= availableHeight) height += heights[start + count++];
            if (count == 0 && start < heights.Length) throw new InvalidOperationException("An editable PowerPoint table row cannot fit with its header. Increase the page height or use Word or Excel.");
            var slide = presentation.AddSlide();
            slide.AddTextBoxPoints(view.Title, view.PageMargin, view.PageMargin, width, 32);
            slide.AddTextBoxPoints(part.Title + (start > 0 ? " (continued)" : ""), view.PageMargin, view.PageMargin + 36, width, 25);
            var table = slide.AddTable(count + 1, part.Headers.Length);
            table.LeftPoints = view.PageMargin; table.TopPoints = view.PageMargin + 72; table.WidthPoints = width; table.HeightPoints = height;
            table.SetColumnWidthsPoints(widths); table.HeaderRow = true;
            var rows = table.RowItems;
            for (int r = 0; r <= count; r++) {
                rows[r].HeightPoints = r == 0 ? headerHeight : heights[start + r - 1];
                for (int c = 0; c < part.Headers.Length; c++) {
                    var cell = rows[r].Cells[c]; cell.Text = r == 0 ? part.Headers[c] : part.Rows[start + r - 1][c];
                    cell.FontSize = 12; cell.FontName = "Arial"; cell.Bold = r == 0; cell.TextAutoFit = PowerPointTextAutoFit.None;
                }
            }
            slide.Notes.Text = "Editable report data. Source revision " + view.ModelRevision + ". " + view.Kind;
            start += count;
        } while (start < part.Rows.Length);
    }
}
