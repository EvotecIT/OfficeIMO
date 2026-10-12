using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            private static void WriteDeclaredRows(TextWriter writer, DirectDataSetSheetModel sheet, int startRow,
                string[] references, DirectColumnWritePlan columns, ExcelTabularStylePlan styles,
                Func<DateTimeOffset, DateTime> offsetStrategy, ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings, CancellationToken ct) {
                var rowWriter = ExcelTabularRowWriter.Create(writer, startRow, sheet.IncludeCellReferences,
                    references, columns.StyleAttributes, columns.ValueStyleColumns, sheet.UseCellValueNumberFormats,
                    offsetStrategy, dateSystem, sharedStrings, styles);
                if (sheet.Table.TryGetObjectRows(out IDirectObjectRows rows)) {
                    rows.WriteRows(rowWriter, ct);
                    return;
                }
                for (int row = 0; row < sheet.Table.RowCount; row++) {
                    ct.ThrowIfCancellationRequested();
                    rowWriter.BeginRow();
                    for (int column = 0; column < sheet.Table.ColumnCount; column++) {
                        WriteDeclaredValue(rowWriter, sheet.Table.GetValue(row, column));
                    }
                    rowWriter.EndRow();
                }
            }

            private static void WriteDeclaredValueRow(ExcelTabularRowWriter writer, object?[] values) {
                writer.BeginRow();
                foreach (object? value in values) WriteDeclaredValue(writer, value);
                writer.EndRow();
            }

            private static void WriteDeclaredValue(ExcelTabularRowWriter writer, object? value) {
                if (value == null || value == DBNull.Value) writer.WriteBlank();
                else writer.Write(value);
            }

            private static void WriteDeclaredColumns(TextWriter writer, double[]? widths, ExcelTabularStylePlan? styles) {
                if (styles == null) { WriteColumns(writer, widths); return; }
                bool opened = false;
                int count = Math.Max(widths?.Length ?? 0, styles.Columns.Length);
                for (int i = 0; i < count; i++) {
                    double width = widths != null && i < widths.Length ? widths[i] : 0;
                    var style = i < styles.Columns.Length ? styles.Columns[i] : null;
                    if (width <= 0 && style == null) continue;
                    if (!opened) { writer.Write("<cols>"); opened = true; }
                    writer.Write("<col min=\""); WriteInvariant(writer, i + 1);
                    writer.Write("\" max=\""); WriteInvariant(writer, i + 1); writer.Write('"');
                    if (width > 0) {
                        writer.Write(" width=\""); WriteColumnWidth(writer, width);
                        writer.Write("\" bestFit=\"1\" customWidth=\"1\"");
                    } else {
                        // Excel loads a styled col with no width as zero-width. Use the same
                        // Calibri-11 fallback as worksheet geometry when no sizing was requested.
                        writer.Write(" width=\"8.43\"");
                    }
                    if (style != null) { writer.Write(" style"); writer.Write(style.Attribute.Substring(2)); }
                    writer.Write("/>");
                }
                if (opened) writer.Write("</cols>");
            }
        }
    }
}
