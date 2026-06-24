using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Resolves the data range for a worksheet or table column by matching its header text.
        /// </summary>
        /// <param name="headerName">Header or table column name to match.</param>
        /// <param name="tableName">Optional table name, display name, or range. When omitted, the worksheet header row is scanned.</param>
        /// <param name="headerRow">Worksheet header row to scan when <paramref name="tableName"/> is omitted. Use 0 to scan the first row of the used range.</param>
        /// <param name="includeHeader">Include the matched header cell in the returned range.</param>
        /// <param name="normalizeHeader">Normalize repeated whitespace before comparing headers.</param>
        /// <returns>A single-column A1 range for the matched column.</returns>
        public string GetColumnRangeByHeader(
            string headerName,
            string? tableName = null,
            int headerRow = 0,
            bool includeHeader = false,
            bool normalizeHeader = true) {
            if (TryGetColumnRangeByHeader(headerName, tableName, headerRow, includeHeader, normalizeHeader, out string range)) {
                return range;
            }

            string scope = string.IsNullOrWhiteSpace(tableName)
                ? $"worksheet '{Name}'"
                : $"table '{tableName}' on worksheet '{Name}'";
            throw new InvalidOperationException($"Column header '{headerName}' was not found in {scope}.");
        }

        /// <summary>
        /// Tries to resolve the data range for a worksheet or table column by matching its header text.
        /// </summary>
        public bool TryGetColumnRangeByHeader(
            string headerName,
            string? tableName,
            int headerRow,
            bool includeHeader,
            bool normalizeHeader,
            out string range) {
            range = string.Empty;
            if (string.IsNullOrWhiteSpace(headerName)) {
                throw new ArgumentNullException(nameof(headerName));
            }

            if (!string.IsNullOrWhiteSpace(tableName)) {
                return TryGetTableColumnRangeByHeader(headerName, tableName!, includeHeader, normalizeHeader, out range);
            }

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                _excelDocument.MaterializeDeferredDataSetImport();
            }

            string usedRange = GetUsedRangeA1();
            if (!A1.TryParseRange(usedRange, out int usedStartRow, out int usedStartColumn, out int usedEndRow, out int usedEndColumn)) {
                return false;
            }

            int effectiveHeaderRow = headerRow > 0 ? headerRow : usedStartRow;
            if (effectiveHeaderRow < usedStartRow || effectiveHeaderRow > usedEndRow) {
                return false;
            }

            for (int column = usedStartColumn; column <= usedEndColumn; column++) {
                if (!TryGetCellText(effectiveHeaderRow, column, out string text)) {
                    continue;
                }

                if (!HeadersMatch(text, headerName, normalizeHeader)) {
                    continue;
                }

                int startRow = includeHeader ? effectiveHeaderRow : effectiveHeaderRow + 1;
                int endRow = Math.Max(startRow, usedEndRow);
                range = BuildSingleColumnRange(startRow, column, endRow);
                return true;
            }

            return false;
        }

        private bool TryGetTableColumnRangeByHeader(
            string headerName,
            string tableName,
            bool includeHeader,
            bool normalizeHeader,
            out string range) {
            range = string.Empty;
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                _excelDocument.MaterializeDeferredDataSetImport();
            }

            Table? table = FindTableByRangeNameOrDisplayName(tableName);
            if (table?.Reference?.Value == null || table.HeaderRowCount?.Value == 0) {
                return false;
            }

            if (!A1.TryParseRange(table.Reference.Value, out int startRow, out int startColumn, out int endRow, out _)) {
                return false;
            }

            int headerRows = (int)(table.HeaderRowCount?.Value ?? 1U);
            int columnOffset = 0;
            foreach (TableColumn tableColumn in table.TableColumns?.Elements<TableColumn>() ?? Enumerable.Empty<TableColumn>()) {
                if (HeadersMatch(tableColumn.Name?.Value, headerName, normalizeHeader)) {
                    int column = startColumn + columnOffset;
                    int firstRow = includeHeader ? startRow : startRow + headerRows;
                    int totalsRows = table.TotalsRowShown?.Value == true
                        ? Math.Max(1, (int)(table.TotalsRowCount?.Value ?? 1U))
                        : 0;
                    int lastRow = Math.Max(firstRow, endRow - totalsRows);
                    range = BuildSingleColumnRange(firstRow, column, lastRow);
                    return true;
                }

                columnOffset++;
            }

            return false;
        }

        private static bool HeadersMatch(string? actual, string expected, bool normalizeHeader) {
            string normalizedActual = ExcelHeaderNameHelper.NormalizeHeader(actual, normalizeHeader);
            string normalizedExpected = ExcelHeaderNameHelper.NormalizeHeader(expected, normalizeHeader);
            return string.Equals(normalizedActual, normalizedExpected, StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildSingleColumnRange(int startRow, int column, int endRow) {
            string start = A1.CellReference(startRow, column);
            string end = A1.CellReference(endRow, column);
            return $"{start}:{end}";
        }
    }
}
