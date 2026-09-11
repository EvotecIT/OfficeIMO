using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Project;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows;

public static partial class ProjectReportWorkflow {
    /// <summary>Creates an editable Word report with native tables. The caller owns and disposes the returned document.</summary>
    public static WordDocument CreateWord(ProjectView view, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view);
        cancellationToken.ThrowIfCancellationRequested();
        var document = WordDocument.Create();
        try {
            document.PageOrientation = view.PageWidth >= view.PageHeight ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;
            document.PageSettings.Width = checked((uint)Math.Round(view.PageWidth * 20));
            document.PageSettings.Height = checked((uint)Math.Round(view.PageHeight * 20));
            document.Margins.Left = document.Margins.Right = checked((uint)Math.Round(view.PageMargin * 20));
            document.Margins.Top = document.Margins.Bottom = checked((int)Math.Round(view.PageMargin * 20));
            var title = document.AddParagraph(view.Title); title.Bold = true; title.FontSize = 20;
            document.AddParagraph(view.Kind.ToString());
            foreach (var part in TableParts(view, cancellationToken)) {
                cancellationToken.ThrowIfCancellationRequested();
                document.AddParagraph(part.Title);
                var table = document.AddTable(part.Rows.Length + 1, part.Headers.Length, WordTableStyle.TableGrid);
                table.SetColumnWidthsPercentage(ColumnWeights(part.Headers));
                table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
                for (int c = 0; c < part.Headers.Length; c++) {
                    var paragraph = table.Rows[0].Cells[c].Paragraphs[0]; paragraph.Text = part.Headers[c]; paragraph.Bold = true;
                }
                for (int r = 0; r < part.Rows.Length; r++) for (int c = 0; c < part.Headers.Length; c++)
                    table.Rows[r + 1].Cells[c].Paragraphs[0].Text = part.Rows[r][c];
                foreach (var row in table.Rows) foreach (var cell in row.Cells) foreach (var paragraph in cell.Paragraphs) {
                    paragraph.FontSize = 10; paragraph.FontFamily = "Arial";
                }
            }
            return document;
        } catch { document.Dispose(); throw; }
    }

    /// <summary>Creates editable PowerPoint table slides. Vector/raster report pages are separate export methods.</summary>
    public static PowerPointPresentation CreatePowerPoint(ProjectView view, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view);
        cancellationToken.ThrowIfCancellationRequested();
        var presentation = PowerPointPresentation.Create();
        try {
            if (view.PageWidth > 4032 || view.PageHeight > 4032) throw new ArgumentOutOfRangeException(nameof(view), "PowerPoint pages cannot exceed 56 inches.");
            presentation.SlideSize.WidthPoints = view.PageWidth; presentation.SlideSize.HeightPoints = view.PageHeight;
            var measurement = new OfficeRasterCanvas(new OfficeRasterImage(1, 1));
            foreach (var part in TableParts(view, cancellationToken)) AddTableSlides(presentation, view, part, measurement, cancellationToken);
            return presentation;
        } catch { presentation.Dispose(); throw; }
    }

    /// <summary>Creates native Excel cells for selected columns and numeric usage buckets, retaining editable values rather than images.</summary>
    public static ExcelDocument CreateExcel(ProjectView view, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view);
        cancellationToken.ThrowIfCancellationRequested();
        var document = ExcelDocument.Create();
        try {
            var sheet = document.AddWorksheet("Report");
            for (int c = 0; c < view.Columns.Count; c++) sheet.CellValue(1, c + 1, ProjectView.ColumnTitle(view.Columns[c]));
            for (int r = 0; r < view.Rows.Count; r++) {
                cancellationToken.ThrowIfCancellationRequested();
                var row = view.Rows[r];
                for (int c = 0; c < view.Columns.Count; c++) {
                    int column = c + 1, line = r + 2;
                    switch (view.Columns[c]) {
                        case ProjectViewColumn.Uid: sheet.CellValue(line, column, row.Uid); break;
                        case ProjectViewColumn.Start when row.Start.HasValue: SetExcelDate(sheet, line, column, row.Start.Value); break;
                        case ProjectViewColumn.Finish when row.Finish.HasValue: SetExcelDate(sheet, line, column, row.Finish.Value); break;
                        case ProjectViewColumn.WorkHours when row.WorkHours.HasValue: sheet.CellValue(line, column, row.WorkHours.Value); break;
                        case ProjectViewColumn.Cost when row.Cost.HasValue: sheet.CellValue(line, column, row.Cost.Value); break;
                        case ProjectViewColumn.PercentComplete when row.PercentComplete.HasValue: sheet.CellValue(line, column, row.PercentComplete.Value); break;
                        default: sheet.CellValue(line, column, row.GetText(view.Columns[c])); break;
                    }
                }
            }
            FinishSheet(sheet, cancellationToken, view.Columns.Count);
            if (IsUsage(view)) {
                var usage = document.AddWorksheet("Work hours"); usage.CellValue(1, 1, "UID"); usage.CellValue(1, 2, "Name");
                for (int b = 0; b < view.Buckets.Count; b++) { usage.CellValue(1, b + 3, view.Buckets[b].Start); usage.FormatCell(1, b + 3, "yyyy-mm-dd"); }
                for (int r = 0; r < view.Rows.Count; r++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    usage.CellValue(r + 2, 1, view.Rows[r].Uid); usage.CellValue(r + 2, 2, view.Rows[r].Name);
                    for (int b = 0; b < view.Buckets.Count; b++) usage.CellValue(r + 2, b + 3, view.Rows[r].BucketWorkHours[b]);
                }
                FinishSheet(usage, cancellationToken, view.Buckets.Count + 2); usage.Freeze(1, 2);
            }
            AddExcelSupplementalTables(document, view, cancellationToken);
            return document;
        } catch { document.Dispose(); throw; }
    }

}
