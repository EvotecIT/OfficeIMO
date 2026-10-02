using System.Threading;
using OfficeIMO.IWork;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    private static IEnumerable<IWorkParagraphStyle> TableParagraphStyles(IWorkTable table) =>
        new[] { table.TextStyles.Body, table.TextStyles.HeaderRow, table.TextStyles.HeaderColumn, table.TextStyles.FooterRow }
            .Concat(table.Cells.Select(cell => cell.ParagraphStyle)).Where(style => style != null).Cast<IWorkParagraphStyle>();

    private static bool HasTableFillDefaults(IWorkTable table) =>
        table.FillStyles.Body != null || table.FillStyles.HeaderRow != null
        || table.FillStyles.HeaderColumn != null || table.FillStyles.FooterRow != null || table.FillStyles.BandedBody != null;

    private static void ApplyTableStyles(ExcelSheet sheet, IWorkTable table, CancellationToken token) {
        if (!TableParagraphStyles(table).Any() && !HasTableFillDefaults(table)) return;
        for (int row = 1; row <= table.RowCount; row++) {
            token.ThrowIfCancellationRequested();
            for (int column = 1; column <= table.ColumnCount; column++) {
                if (table.GetFill(row, column)?.Color is { } fill) sheet.CellAt(row, column).SetFillColor(fill.RgbHex);
                if (table.GetParagraphStyle(row, column) is not { } style) continue;
                ExcelCell cell = sheet.CellAt(row, column);
                IWorkTextStyle text = style.TextStyle;
                if (text.Bold.HasValue) cell.SetBold(text.Bold.Value);
                if (text.Italic.HasValue) sheet.CellItalic(row, column, text.Italic.Value);
                if (text.Underline.HasValue) sheet.CellUnderline(row, column, text.Underline.Value);
                if (text.Strikethrough.HasValue) sheet.CellStrikethrough(row, column, text.Strikethrough.Value);
                if (text.FontSizePoints.HasValue) cell.SetFontSize(text.FontSizePoints.Value);
                if (!string.IsNullOrWhiteSpace(text.FontName)) cell.SetFontName(text.FontName!);
                if (text.Color != null) cell.SetFontColor(text.Color.RgbHex);
                if (style.Alignment is { } alignment) sheet.CellAlign(row, column, alignment switch {
                    IWorkTextAlignment.Center => ExcelHorizontalAlignment.Center,
                    IWorkTextAlignment.Right => ExcelHorizontalAlignment.Right,
                    IWorkTextAlignment.Justified => ExcelHorizontalAlignment.Justify,
                    IWorkTextAlignment.Natural => OfficeTextElements.ResolveBaseDirection(table.GetCell(row, column)?.CachedDisplayText ?? string.Empty)
                        == OfficeTextDirection.RightToLeft ? ExcelHorizontalAlignment.Right : ExcelHorizontalAlignment.Left,
                    _ => ExcelHorizontalAlignment.Left
                });
            }
        }
    }

    private static bool HasUnsupportedTableTextStyle(IWorkParagraphStyle style) =>
        style.TextStyle.BackgroundColor != null || style.TextStyle.Color is { Alpha: < byte.MaxValue }
        || style.FirstLineIndentPoints.GetValueOrDefault() != 0 || style.LeftIndentPoints.GetValueOrDefault() != 0
        || style.RightIndentPoints.GetValueOrDefault() != 0 || style.SpaceBeforePoints.GetValueOrDefault() != 0
        || style.TabStops?.Count > 0 || style.LineSpacingMultiplier.HasValue || style.SpaceAfterPoints.GetValueOrDefault() != 0 || style.PageBreakBefore == true
        || style.KeepWithNext == true || style.KeepLinesTogether == true;
}
