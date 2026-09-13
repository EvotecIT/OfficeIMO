using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Project;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows;

public static partial class ProjectReportWorkflow {
    /// <summary>Creates chart pages and editable Word tables. The caller owns and disposes the returned document.</summary>
    public static WordDocument CreateWord(ProjectView view, CancellationToken cancellationToken = default) {
        return CreateWord(view, new ProjectOfficeReportOptions(), cancellationToken);
    }

    /// <summary>Creates a Word report with explicit chart and table selection.</summary>
    public static WordDocument CreateWord(ProjectView view, ProjectOfficeReportOptions options, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view);
        ValidateOfficeOptions(options);
        cancellationToken.ThrowIfCancellationRequested();
        var document = WordDocument.Create();
        try {
            document.PageOrientation = view.PageWidth >= view.PageHeight ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;
            document.PageSettings.Width = checked((uint)Math.Round(view.PageWidth * 20));
            document.PageSettings.Height = checked((uint)Math.Round(view.PageHeight * 20));
            document.Margins.Left = document.Margins.Right = checked((uint)Math.Round(view.PageMargin * 20));
            document.Margins.Top = document.Margins.Bottom = checked((int)Math.Round(view.PageMargin * 20));
            if (options.IncludeCharts && view.Kind != ProjectViewKind.Table) AddWordCharts(document, view, options, cancellationToken);
            if (!options.IncludeDataTables && view.Kind != ProjectViewKind.Table) return document;
            var title = document.AddParagraph(view.Title); title.Bold = true; title.FontSize = 20; title.FontFamily = "Arial"; title.ColorHex = "183047";
            foreach (var part in TableParts(view, cancellationToken)) {
                cancellationToken.ThrowIfCancellationRequested();
                var heading = document.AddParagraph(part.Title); heading.FontFamily = "Arial"; heading.FontSize = 13;
                heading.Bold = true; heading.ColorHex = "183047"; heading.KeepWithNext = true;
                heading.LineSpacingBeforePoints = 10; heading.LineSpacingAfterPoints = 6;
                var table = document.AddTable(part.Rows.Length + 1, part.Headers.Length, WordTableStyle.TableGrid);
                table.SetColumnWidthsPercentage(ColumnWeights(part.Headers));
                table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Single, 4, OfficeColor.ParseHex("#E3EAF0"));
                var tableRows = table.Rows;
                tableRows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
                for (int r = 0; r < tableRows.Count; r++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var cells = tableRows[r].Cells;
                    for (int c = 0; c < cells.Count; c++) {
                        var cell = cells[c];
                        cell.ShadingFillColorHex = r == 0 ? "183047" : r % 2 == 1 ? "F4F7FA" : "FFFFFF";
                        cell.MarginTopWidth = cell.MarginBottomWidth = 80;
                        cell.MarginLeftWidth = cell.MarginRightWidth = 120;
                        var paragraph = cell.Paragraphs[0];
                        paragraph.Text = r == 0 ? part.Headers[c] : part.Rows[r - 1][c]; paragraph.Bold = r == 0;
                        paragraph.FontSize = 11; paragraph.FontFamily = "Arial"; paragraph.ColorHex = r == 0 ? "FFFFFF" : "183047";
                    }
                }
            }
            return document;
        } catch { document.Dispose(); throw; }
    }

    /// <summary>Creates chart slides and editable PowerPoint table slides.</summary>
    public static PowerPointPresentation CreatePowerPoint(ProjectView view, CancellationToken cancellationToken = default) {
        return CreatePowerPoint(view, new ProjectOfficeReportOptions(), cancellationToken);
    }

    /// <summary>Creates a presentation with explicit chart and table selection.</summary>
    public static PowerPointPresentation CreatePowerPoint(ProjectView view, ProjectOfficeReportOptions options, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view);
        ValidateOfficeOptions(options);
        cancellationToken.ThrowIfCancellationRequested();
        var presentation = PowerPointPresentation.Create();
        try {
            if (view.PageWidth > 4032 || view.PageHeight > 4032) throw new ArgumentOutOfRangeException(nameof(view), "PowerPoint pages cannot exceed 56 inches.");
            presentation.SlideSize.WidthPoints = view.PageWidth; presentation.SlideSize.HeightPoints = view.PageHeight;
            if (options.IncludeCharts && view.Kind != ProjectViewKind.Table) AddPowerPointCharts(presentation, view, options, cancellationToken);
            if (!options.IncludeDataTables && view.Kind != ProjectViewKind.Table) return presentation;
            var measurement = new OfficeRasterCanvas(new OfficeRasterImage(1, 1));
            foreach (var part in TableParts(view, cancellationToken)) AddTableSlides(presentation, view, part, measurement, cancellationToken);
            return presentation;
        } catch { presentation.Dispose(); throw; }
    }

    /// <summary>Creates native Excel cells for selected columns and numeric usage buckets, retaining editable values rather than images.</summary>
    public static ExcelDocument CreateExcel(ProjectView view, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view);
        cancellationToken.ThrowIfCancellationRequested();
        ValidateExcelGrid(view);
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
            FinishSheet(sheet, cancellationToken, view.Rows.Count + 1, view.Columns.Count);
            if (IsUsage(view)) {
                var usage = document.AddWorksheet("Work hours"); usage.CellValue(1, 1, "UID"); usage.CellValue(1, 2, "Name");
                for (int b = 0; b < view.Buckets.Count; b++) { usage.CellValue(1, b + 3, view.Buckets[b].Start); usage.FormatCell(1, b + 3, "yyyy-mm-dd"); }
                for (int r = 0; r < view.Rows.Count; r++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    usage.CellValue(r + 2, 1, view.Rows[r].Uid); usage.CellValue(r + 2, 2, view.Rows[r].Name);
                    for (int b = 0; b < view.Buckets.Count; b++) usage.CellValue(r + 2, b + 3, view.Rows[r].BucketWorkHours[b]);
                }
                FinishSheet(usage, cancellationToken, view.Rows.Count + 1, view.Buckets.Count + 2, 2); usage.Freeze(1, 2);
            }
            AddExcelSupplementalTables(document, view, cancellationToken);
            foreach (var reportSheet in document.Sheets) {
                reportSheet.SetHeaderFooter(headerLeft: view.Title.Replace("&", "&&"), headerRight: "&A", footerRight: "Page &P of &N");
            }
            return document;
        } catch { document.Dispose(); throw; }
    }

}
