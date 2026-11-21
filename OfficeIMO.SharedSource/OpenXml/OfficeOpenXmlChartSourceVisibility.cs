using System.Globalization;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>Source-cell rules for charts that plot visible cells only (c:plotVisOnly).</summary>
    /// <remarks>Office plots each referenced range without its hidden cells and closes the gaps, so a
    /// plotted point is the n-th visible cell of its range. A series whose value cells are all hidden is
    /// left out. Hidden series-name cells do not hide a series.</remarks>
    internal static class OfficeOpenXmlChartSourceVisibility {
        /// <summary>Office plots visible cells only unless c:plotVisOnly is explicitly off, including when it is absent.</summary>
        internal static bool PlotsVisibleCellsOnly(C.Chart? chart) =>
            chart?.GetFirstChild<C.PlotVisibleOnly>()?.Val?.Value != false;

        /// <summary>Parses a single-area, sheet-qualified A1 reference such as <c>'My Sheet'!$A$2:$A$5</c>.</summary>
        internal static bool TryParseRange(string? formula, out string sheetName,
            out uint firstRow, out uint lastRow, out uint firstColumn, out uint lastColumn) {
            sheetName = string.Empty;
            firstRow = lastRow = firstColumn = lastColumn = 0;
            if (string.IsNullOrWhiteSpace(formula)) return false;
            string text = formula!.Trim();
            int separator = text.LastIndexOf('!');
            if (separator <= 0 || separator >= text.Length - 1) return false;
            string sheet = text.Substring(0, separator).Trim();
            if (sheet.IndexOf('[') >= 0 || sheet.IndexOf(']') >= 0) return false;
            if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'')
                sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
            else if (sheet.IndexOf('\'') >= 0) return false;
            if (sheet.Length == 0) return false;
            string range = text.Substring(separator + 1).Trim();
            string[] endpoints = range.Split(':');
            if (endpoints.Length is < 1 or > 2 ||
                !TryParseCell(endpoints[0], out firstRow, out firstColumn) ||
                !TryParseCell(endpoints[endpoints.Length - 1], out lastRow, out lastColumn) ||
                lastRow < firstRow || lastColumn < firstColumn) return false;
            sheetName = sheet;
            return true;
        }

        private static bool TryParseCell(string cell, out uint row, out uint column) {
            row = column = 0;
            string reference = cell.Trim().Replace("$", string.Empty).ToUpperInvariant();
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

    /// <summary>The rows and columns of one worksheet that a visible-cells-only chart leaves out.</summary>
    /// <remarks>A row is hidden when its record says so. With a zero-height sheet default
    /// (<c>sheetFormatPr zeroHeight="1"</c>) every row without a record is hidden too, while a row record
    /// that does not set <c>hidden="1"</c> keeps its row visible.</remarks>
    internal sealed class OfficeOpenXmlChartHiddenCells {
        private readonly bool _rowsHiddenByDefault;
        // Sorted row numbers: the hidden rows, or with a zero-height default the rows that stay visible.
        private readonly uint[] _rows;
        private readonly int[] _columnPrefix = new int[16385];

        internal OfficeOpenXmlChartHiddenCells(S.Worksheet worksheet) {
            // Attribute text is read leniently, so malformed or numerous plain row records cost nothing and never throw.
            _rowsHiddenByDefault = IsTrue(worksheet.GetFirstChild<S.SheetFormatProperties>()?.ZeroHeight);
            var rows = new List<uint>();
            foreach (S.Row row in worksheet.GetFirstChild<S.SheetData>()?.Elements<S.Row>() ?? Enumerable.Empty<S.Row>()) {
                if (IsTrue(row.Hidden) == _rowsHiddenByDefault) continue;
                if (uint.TryParse(row.RowIndex?.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out uint index)) rows.Add(index);
            }
            rows.Sort();
            _rows = rows.Distinct().ToArray();
            var columnDelta = new int[16386];
            foreach (S.Column column in worksheet.Elements<S.Columns>().SelectMany(columns => columns.Elements<S.Column>())) {
                if (!IsTrue(column.Hidden) ||
                    !uint.TryParse(column.Min?.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out uint minimum) ||
                    !uint.TryParse(column.Max?.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out uint maximum) ||
                    minimum < 1 || minimum > maximum || minimum > 16384) continue;
                columnDelta[minimum]++;
                columnDelta[Math.Min(maximum, 16384U) + 1]--;
            }
            int depth = 0;
            for (int column = 1; column < _columnPrefix.Length; column++) {
                depth += columnDelta[column];
                _columnPrefix[column] = _columnPrefix[column - 1] + (depth > 0 ? 1 : 0);
            }
        }

        private static bool IsTrue(DocumentFormat.OpenXml.BooleanValue? value) => value?.InnerText is "1" or "true";

        /// <summary>Whether any row or column of the worksheet is hidden.</summary>
        internal bool HasHiddenCells => _rowsHiddenByDefault || _rows.Length > 0 || _columnPrefix[16384] > 0;

        internal bool IsRowHidden(uint row) => (Array.BinarySearch(_rows, row) >= 0) != _rowsHiddenByDefault;

        internal bool IsColumnHidden(uint column) => column is >= 1 and <= 16384 && _columnPrefix[column] > _columnPrefix[column - 1];

        /// <summary>Whether the range contains a hidden row or column.</summary>
        internal bool Intersects(uint firstRow, uint lastRow, uint firstColumn, uint lastColumn) {
            if (_columnPrefix[Math.Min(lastColumn, 16384U)] > _columnPrefix[firstColumn - 1]) return true;
            long listed = CountListedRows(firstRow, lastRow);
            return _rowsHiddenByDefault ? listed < (long)lastRow - firstRow + 1 : listed > 0;
        }

        /// <summary>The zero-based offsets of the visible cells of a range in row-major order, or null when all are visible.</summary>
        internal int[]? GetVisibleOffsets(uint firstRow, uint lastRow, uint firstColumn, uint lastColumn) {
            if (!Intersects(firstRow, lastRow, firstColumn, lastColumn)) return null;
            var visible = new List<int>();
            int offset = 0;
            for (uint row = firstRow; row <= lastRow; row++) {
                bool rowHidden = IsRowHidden(row);
                for (uint column = firstColumn; column <= lastColumn; column++, offset++) {
                    if (!rowHidden && !IsColumnHidden(column)) visible.Add(offset);
                }
            }
            return visible.ToArray();
        }

        private long CountListedRows(uint firstRow, uint lastRow) {
            int start = Array.BinarySearch(_rows, firstRow);
            if (start < 0) start = ~start;
            int end = Array.BinarySearch(_rows, lastRow);
            end = end < 0 ? ~end : end + 1;
            return Math.Max(0, end - start);
        }
    }
}
