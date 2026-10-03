using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        private static bool HasUnsupportedBubbleSourceVisibility(ChartPart part, C.Chart chart) {
            var visibleOnly = chart.GetFirstChild<C.PlotVisibleOnly>();
            if (visibleOnly == null || visibleOnly.Val?.Value == false) return false;
            // Literal cached charts have no workbook source whose hidden rows can affect plotting.
            if (chart.PlotArea?.Descendants<C.Formula>().Any(formula => !string.IsNullOrWhiteSpace(formula.Text)) != true) return false;
            try {
                var embedded = OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(part);
                if (embedded == null) return true;
                using var stream = embedded.GetStream(FileMode.Open, FileAccess.Read);
                byte[] bytes = OfficeOpenXmlChartWorkbookSecurity.ReadAndValidate(stream);
                using var workbookStream = new MemoryStream(bytes, writable: false);
                using var workbook = SpreadsheetDocument.Open(workbookStream, false, OfficeOpenXmlChartWorkbookSecurity.CreateOpenSettings());
                if (workbook.WorkbookPart == null) return true;
                bool anyHidden = workbook.WorkbookPart.WorksheetParts.Any(sheet =>
                    sheet.Worksheet?.Descendants<S.Row>().Any(row => row.Hidden?.Value == true) == true ||
                    sheet.Worksheet?.Descendants<S.Column>().Any(column => column.Hidden?.Value == true) == true);
                if (!anyHidden) return false;
                var sheets = new Dictionary<string, WorksheetPart>(StringComparer.OrdinalIgnoreCase);
                foreach (S.Sheet sheet in workbook.WorkbookPart.Workbook?.Sheets?.Elements<S.Sheet>() ?? Enumerable.Empty<S.Sheet>()) {
                    if (sheet.Name?.Value == null || sheet.Id?.Value == null ||
                        workbook.WorkbookPart.GetPartById(sheet.Id.Value) is not WorksheetPart worksheetPart ||
                        worksheetPart.Worksheet == null || sheets.ContainsKey(sheet.Name.Value)) return true;
                    sheets.Add(sheet.Name.Value, worksheetPart);
                }
                var hiddenBySheet = new Dictionary<WorksheetPart, HiddenWorksheet>();
                C.Formula[] formulas = chart.PlotArea!.Descendants<C.Formula>().Take(10001).ToArray();
                if (formulas.Length > 10000) return true;
                foreach (C.Formula formula in formulas) {
                    if (string.IsNullOrWhiteSpace(formula.Text)) continue;
                    if (!TryParseChartRange(formula.Text, out string sheetName,
                        out uint firstRow, out uint lastRow, out uint firstColumn, out uint lastColumn)) return true;
                    if (!sheets.TryGetValue(sheetName, out WorksheetPart? worksheetPart)) return true;
                    if (!hiddenBySheet.TryGetValue(worksheetPart, out HiddenWorksheet? hidden)) {
                        hidden = new HiddenWorksheet(worksheetPart.Worksheet!);
                        hiddenBySheet.Add(worksheetPart, hidden);
                    }
                    if (hidden.Intersects(firstRow, lastRow, firstColumn, lastColumn))
                        return true;
                }
                return false;
            } catch {
                return true;
            }
        }

        private sealed class HiddenWorksheet {
            private readonly uint[] _rows;
            private readonly int[] _columnPrefix = new int[16385];

            internal HiddenWorksheet(S.Worksheet worksheet) {
                _rows = worksheet.Descendants<S.Row>()
                    .Where(row => row.Hidden?.Value == true && row.RowIndex?.Value is uint)
                    .Select(row => row.RowIndex!.Value).OrderBy(index => index).ToArray();
                var columnDelta = new int[16386];
                foreach (S.Column column in worksheet.Descendants<S.Column>().Where(column => column.Hidden?.Value == true)) {
                    if (column.Min?.Value is not uint minimum || column.Max?.Value is not uint maximum ||
                        minimum < 1 || maximum > 16384 || minimum > maximum) continue;
                    columnDelta[minimum]++;
                    columnDelta[maximum + 1]--;
                }
                int depth = 0;
                for (int column = 1; column < _columnPrefix.Length; column++) {
                    depth += columnDelta[column];
                    _columnPrefix[column] = _columnPrefix[column - 1] + (depth > 0 ? 1 : 0);
                }
            }

            internal bool Intersects(uint firstRow, uint lastRow, uint firstColumn, uint lastColumn) {
                int rowIndex = Array.BinarySearch(_rows, firstRow);
                if (rowIndex < 0) rowIndex = ~rowIndex;
                return rowIndex < _rows.Length && _rows[rowIndex] <= lastRow ||
                    _columnPrefix[lastColumn] > _columnPrefix[firstColumn - 1];
            }
        }

        private static bool TryParseChartRange(string formula, out string sheetName,
            out uint firstRow, out uint lastRow, out uint firstColumn, out uint lastColumn) {
            sheetName = string.Empty;
            firstRow = lastRow = firstColumn = lastColumn = 0;
            int separator = formula.LastIndexOf('!');
            if (separator <= 0 || separator >= formula.Length - 1) return false;
            string sheet = formula.Substring(0, separator).Trim();
            if (sheet.IndexOf('[') >= 0 || sheet.IndexOf(']') >= 0) return false;
            if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'')
                sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
            else if (sheet.IndexOf('\'') >= 0) return false;
            if (sheet.Length == 0) return false;
            string range = formula.Substring(separator + 1).Trim();
            string[] endpoints = range.Split(':');
            if (endpoints.Length is < 1 or > 2 ||
                !TryParseChartCell(endpoints[0], out firstRow, out firstColumn) ||
                !TryParseChartCell(endpoints[endpoints.Length - 1], out lastRow, out lastColumn) ||
                lastRow < firstRow || lastColumn < firstColumn) return false;
            sheetName = sheet;
            return true;
        }

        private static bool TryParseChartCell(string cell, out uint row, out uint column) {
            row = column = 0;
            string reference = cell.Replace("$", string.Empty).ToUpperInvariant();
            int index = 0;
            while (index < reference.Length && reference[index] >= 'A' && reference[index] <= 'Z') {
                column = column * 26U + (uint)(reference[index] - 'A' + 1);
                if (column > 16384U) return false;
                index++;
            }
            return column > 0 && index < reference.Length &&
                uint.TryParse(reference.Substring(index), NumberStyles.None, CultureInfo.InvariantCulture, out row) &&
                row is >= 1 and <= 1048576;
        }
    }
}
