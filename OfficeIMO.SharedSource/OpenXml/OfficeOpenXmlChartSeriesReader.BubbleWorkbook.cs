using System;
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
                C.Formula[] formulas = chart.PlotArea!.Descendants<C.Formula>().Take(10001).ToArray();
                if (formulas.Length > 10000) return true;
                foreach (C.Formula formula in formulas) {
                    if (string.IsNullOrWhiteSpace(formula.Text)) continue;
                    if (!TryParseChartRange(formula.Text, out string sheetName,
                        out uint firstRow, out uint lastRow, out uint firstColumn, out uint lastColumn)) return true;
                    S.Sheet? sheet = workbook.WorkbookPart.Workbook?.Sheets?.Elements<S.Sheet>()
                        .FirstOrDefault(item => string.Equals(item.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
                    if (sheet?.Id?.Value == null ||
                        workbook.WorkbookPart.GetPartById(sheet.Id.Value) is not WorksheetPart worksheetPart ||
                        worksheetPart.Worksheet == null) return true;
                    if (worksheetPart.Worksheet.Descendants<S.Row>().Any(row =>
                            row.Hidden?.Value == true && row.RowIndex?.Value is uint index &&
                            index >= firstRow && index <= lastRow) ||
                        worksheetPart.Worksheet.Descendants<S.Column>().Any(column =>
                            column.Hidden?.Value == true && column.Min?.Value is uint minimum &&
                            column.Max?.Value is uint maximum && minimum <= lastColumn && maximum >= firstColumn))
                        return true;
                }
                return false;
            } catch {
                return true;
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
