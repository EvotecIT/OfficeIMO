using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Returns a lightweight object wrapper for a single cell.
        /// </summary>
        public ExcelCell CellAt(int row, int column) => new ExcelCell(this, row, column);

        /// <summary>
        /// Returns a lightweight object wrapper for an A1 range.
        /// </summary>
        public ExcelRange Range(string a1Range) => new ExcelRange(this, a1Range);

        /// <summary>
        /// Returns a lightweight object wrapper for a table by name, display name, or range.
        /// </summary>
        public ExcelTable Table(string nameOrRange) => new ExcelTable(this, nameOrRange);

        internal ExcelCellData GetCellValueSnapshot(int row, int column) {
            var cell = TryGetExistingCell(row, column);
            return GetCellValueSnapshot(cell);
        }

        private ExcelCellData GetCellValueSnapshot(Cell? cell) {
            if (cell == null) {
                return new ExcelCellData(ExcelCellDataKind.Blank, null);
            }

            string? cached = cell.CellValue?.Text;
            if (cell.CellFormula != null) {
                object? formulaValue = double.TryParse(cached, NumberStyles.Float, CultureInfo.InvariantCulture, out double cachedNumber)
                    ? cachedNumber
                    : cached;
                return new ExcelCellData(ExcelCellDataKind.Formula, formulaValue, cell.CellFormula.Text, cached);
            }

            if (cell.CellValue == null && cell.InlineString == null) {
                return new ExcelCellData(ExcelCellDataKind.Blank, null);
            }

            var type = cell.DataType?.Value;
            if (type == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean) {
                return new ExcelCellData(ExcelCellDataKind.Boolean, cached == "1" || string.Equals(cached, "true", StringComparison.OrdinalIgnoreCase), cachedText: cached);
            }

            if (type == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error) {
                return new ExcelCellData(ExcelCellDataKind.Error, cached, cachedText: cached);
            }

            string text = GetCellText(cell);
            if (type == DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                || type == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                || type == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                return new ExcelCellData(ExcelCellDataKind.Text, text, cachedText: text);
            }

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                return new ExcelCellData(ExcelCellDataKind.Number, number, cachedText: text);
            }

            return string.IsNullOrEmpty(text)
                ? new ExcelCellData(ExcelCellDataKind.Blank, null)
                : new ExcelCellData(ExcelCellDataKind.Text, text, cachedText: text);
        }

        internal string GetCellFormattedText(int row, int column, IFormatProvider? provider) {
            var value = GetCellValueSnapshot(row, column);
            if (value.Value is IFormattable formattable) {
                return formattable.ToString(null, provider ?? CultureInfo.CurrentCulture) ?? string.Empty;
            }

            return Convert.ToString(value.Value, provider as CultureInfo ?? CultureInfo.CurrentCulture) ?? value.CachedText ?? string.Empty;
        }

        /// <summary>
        /// Applies a number format to every cell in the range.
        /// </summary>
        public void FormatRange(string a1Range, string numberFormat) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLock(() => FormatRangeCore(r1, c1, r2, c2, numberFormat));
        }

        /// <summary>
        /// Applies a solid fill to every cell in the range.
        /// </summary>
        public void FillRange(string a1Range, string hexColor) {
            if (string.IsNullOrWhiteSpace(hexColor)) return;
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLock(() => FillRangeCore(r1, c1, r2, c2, hexColor));
        }

        /// <summary>
        /// Clears selected parts of every cell and attached worksheet metadata in the range.
        /// </summary>
        public void ClearRange(string a1Range, ExcelClearOptions options = ExcelClearOptions.All) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (options == ExcelClearOptions.None) {
                return;
            }

            bool clearCellFields = options.HasFlag(ExcelClearOptions.Values)
                || options.HasFlag(ExcelClearOptions.Formulas)
                || options.HasFlag(ExcelClearOptions.Styles);
            if (clearCellFields) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLock(() => {
                var ws = WorksheetRoot;
                bool worksheetChanged = false;

                if (clearCellFields) {
                    worksheetChanged |= ClearExistingCellFieldsInRange((r1, c1, r2, c2), options);
                }

                if (options.HasFlag(ExcelClearOptions.Comments)) {
                    worksheetChanged |= ClearCommentsInRange(r1, c1, r2, c2);
                }

                if (options.HasFlag(ExcelClearOptions.Hyperlinks)) {
                    worksheetChanged |= ClearHyperlinksInRange(ws, (r1, c1, r2, c2));
                }

                if (options.HasFlag(ExcelClearOptions.DataValidations)) {
                    RemoveDataValidationsCore(a1Range);
                }

                if (options.HasFlag(ExcelClearOptions.ConditionalFormatting)) {
                    ClearConditionalFormattingCore(a1Range);
                }

                if (options.HasFlag(ExcelClearOptions.Merges)) {
                    UnmergeRangeCore((r1, c1, r2, c2));
                }

                if (options.HasFlag(ExcelClearOptions.Sparklines)) {
                    worksheetChanged |= ClearSparklinesInRange((r1, c1, r2, c2));
                }

                if (worksheetChanged) {
                    ws.Save();
                    ClearHeaderCache();
                }
            });
        }

        private bool ClearExistingCellFieldsInRange((int r1, int c1, int r2, int c2) bounds, ExcelClearOptions options) {
            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return false;
            }

            bool clearValues = options.HasFlag(ExcelClearOptions.Values);
            bool clearFormulas = options.HasFlag(ExcelClearOptions.Formulas);
            bool clearStyles = options.HasFlag(ExcelClearOptions.Styles);
            bool changed = false;

            foreach (var row in sheetData.Elements<Row>()) {
                uint rowIndex = row.RowIndex?.Value ?? 0U;
                if (rowIndex < (uint)bounds.r1 || rowIndex > (uint)bounds.r2) {
                    continue;
                }

                foreach (var cell in row.Elements<Cell>()) {
                    if (cell.CellReference?.Value is not string reference) {
                        continue;
                    }

                    int columnIndex = GetColumnIndex(reference);
                    if (columnIndex < bounds.c1 || columnIndex > bounds.c2) {
                        continue;
                    }

                    if (clearValues && (cell.CellValue != null || cell.DataType != null || cell.InlineString != null)) {
                        cell.CellValue = null;
                        cell.DataType = null;
                        cell.InlineString = null;
                        changed = true;
                    }

                    if (clearFormulas && cell.CellFormula != null) {
                        cell.CellFormula = null;
                        changed = true;
                    }

                    if (clearStyles && cell.StyleIndex != null) {
                        cell.StyleIndex = null;
                        changed = true;
                    }
                }
            }

            return changed;
        }
    }
}
