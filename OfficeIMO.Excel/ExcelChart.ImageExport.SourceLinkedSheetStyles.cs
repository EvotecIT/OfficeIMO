using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

public sealed partial class ExcelChart {
    private sealed class SourceLinkedSheetStyles {
        internal readonly Dictionary<string, uint?> Cells = new(StringComparer.OrdinalIgnoreCase);
        internal readonly Dictionary<int, uint> Rows = new();
        internal readonly Dictionary<int, uint> Columns = new();
        internal uint Resolve(int row, int column, string reference) =>
            Cells.TryGetValue(reference, out uint? cellStyle) && cellStyle.HasValue ? cellStyle.Value :
            Rows.TryGetValue(row, out uint rowStyle) ? rowStyle :
            Columns.TryGetValue(column, out uint columnStyle) ? columnStyle : 0;
    }

    private static SourceLinkedSheetStyles? ReadSourceLinkedSheetStyles(ExcelSheet sheet, ref int remainingRecords) {
        var result = new SourceLinkedSheetStyles();
        var worksheet = sheet.WorksheetPart.Worksheet;
        if (worksheet == null) return null;
        foreach (var columns in worksheet.Elements<S.Columns>()) {
            if (--remainingRecords < 0) return null;
            foreach (var column in columns.Elements<S.Column>()) {
                if (--remainingRecords < 0) return null;
                if (column.Style == null) continue;
                if (!TryStyleIndex(column.Style.InnerText, out uint style) ||
                    !uint.TryParse(column.Min?.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out uint first) ||
                    !uint.TryParse(column.Max?.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out uint last) ||
                    first == 0 || first > last || last > 16384 || last - first + 1 > remainingRecords) return null;
                remainingRecords -= (int)(last - first + 1);
                for (uint index = first; index <= last; index++) {
                    if (result.Columns.ContainsKey((int)index)) return null;
                    result.Columns.Add((int)index, style);
                }
            }
        }
        var data = worksheet.GetFirstChild<S.SheetData>();
        if (data == null) return null;
        int previousRow = 0;
        foreach (var row in data.Elements<S.Row>()) {
            if (--remainingRecords < 0) return null;
            int rowIndex;
            if (row.RowIndex != null &&
                !uint.TryParse(row.RowIndex.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out _)) {
                if (row.HasChildren || row.StyleIndex != null || row.CustomFormat != null) return null;
                continue;
            }
            if (row.RowIndex?.Value is uint declaredRow) {
                if (declaredRow == 0 || declaredRow > 1048576) return null;
                rowIndex = (int)declaredRow;
            } else {
                if (previousRow >= 1048576) return null;
                rowIndex = previousRow + 1;
            }
            previousRow = rowIndex;
            string? customFormat = row.CustomFormat?.InnerText;
            if (customFormat != null && customFormat != "true" && customFormat != "1" && customFormat != "false" && customFormat != "0") return null;
            if (customFormat == "true" || customFormat == "1") {
                if (!TryStyleIndex(row.StyleIndex?.InnerText, out uint rowStyle) ||
                    result.Rows.ContainsKey(rowIndex)) return null;
                result.Rows.Add(rowIndex, rowStyle);
            }
            int previousColumn = 0;
            foreach (var cell in row.Elements<S.Cell>()) {
                if (--remainingRecords < 0) return null;
                string? reference = cell.CellReference?.Value;
                if (reference == null) {
                    if (previousColumn >= 16384) return null;
                    previousColumn++;
                    reference = A1.ColumnIndexToLetters(previousColumn) + rowIndex.ToString(CultureInfo.InvariantCulture);
                } else {
                    if (!A1.TryParseCellReferenceFast(reference, out int cellRow, out int column) ||
                        cellRow != rowIndex || column < 1 || column > 16384) return null;
                    previousColumn = column;
                }
                uint? style = null;
                if (cell.StyleIndex != null) {
                    if (!TryStyleIndex(cell.StyleIndex.InnerText, out uint explicitStyle)) return null;
                    style = explicitStyle;
                }
                if (result.Cells.ContainsKey(reference)) return null;
                result.Cells.Add(reference, style);
            }
        }
        return result;
    }

    private static bool TryStyleIndex(string? text, out uint style) =>
        uint.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out style);
}
