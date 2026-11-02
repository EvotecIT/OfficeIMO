using System.Globalization;
using OfficeIMO.Spreadsheet;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Returns the cell at a 1-based row and column. The handle is a view over the worksheet and stores no cell
        /// state, so creating one allocates nothing and does not create the cell until a value or style is written.
        /// </summary>
        /// <exception cref="ArgumentOutOfRangeException">The row or column is outside the worksheet.</exception>
        public ExcelCell Cell(int row, int column) {
            if (row < 1 || row > ExcelA1.MaxRows) throw new ArgumentOutOfRangeException(nameof(row));
            if (column < 1 || column > ExcelA1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(column));
            return new ExcelCell(this, row, column);
        }

        /// <summary>Returns the cell at an A1 reference on this worksheet, such as <c>B7</c> or <c>$B$7</c>.</summary>
        /// <exception cref="ArgumentException">The reference is not one cell on this worksheet.</exception>
        public ExcelCell Cell(string reference) {
            if (string.IsNullOrWhiteSpace(reference)) throw new ArgumentNullException(nameof(reference));
            var syntax = SpreadsheetRangeReference.Parse(reference, SpreadsheetAddressDialect.ExcelA1);
            if (syntax.End != null || !syntax.Start.IsCell || syntax.Start.SheetName != null)
                throw new ArgumentException("Cell reference must identify one cell on the current worksheet.", nameof(reference));
            return Cell(checked((int)syntax.Start.Row!.Value), syntax.Start.Column!.Value);
        }

        /// <summary>
        /// Returns a range by A1 address, such as <c>A1:V1</c>. A single cell reference returns a one-cell range.
        /// </summary>
        public ExcelRange Range(string a1Range) => new ExcelRange(this, a1Range);

        /// <summary>Returns the range between two 1-based corner cells.</summary>
        /// <exception cref="ArgumentOutOfRangeException">A corner is outside the worksheet or the corners are reversed.</exception>
        public ExcelRange Range(int firstRow, int firstColumn, int lastRow, int lastColumn) =>
            new ExcelRange(this, firstRow, firstColumn, lastRow, lastColumn);

        /// <summary>
        /// Returns a lightweight object wrapper for a table by name, display name, or range.
        /// </summary>
        public ExcelTable Table(string nameOrRange) => new ExcelTable(this, nameOrRange);

        internal ExcelCellData GetCellValueSnapshot(int row, int column) {
            MaterializeDeferredDataSetImportIfNeeded();
            Cell? cell = ProbeCell(row, column, out ExcelStoredRow? storedRow, out int storedIndex);
            return storedRow != null ? GetStoredCellValueSnapshot(storedRow.CellAt(storedIndex)) : GetCellValueSnapshot(cell);
        }

        private ExcelCellData GetCellValueSnapshot(Cell? cell) {
            if (cell == null) {
                return new ExcelCellData(ExcelCellDataKind.Blank, null);
            }

            string? cached = ResolveRichValueError(cell, ExcelWorksheetScan.FindValue(cell)?.Text);
            var type = cell.DataType?.Value;
            if (ExcelWorksheetScan.FindFormula(cell) != null) {
                object? formulaValue = cached;
                if (cached != null) {
                    if (type == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean) {
                        formulaValue = cached == "1" || string.Equals(cached, "true", StringComparison.OrdinalIgnoreCase);
                    } else if (type == DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                        || type == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                        || type == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                        formulaValue = GetCellText(cell);
                    } else if (type != DocumentFormat.OpenXml.Spreadsheet.CellValues.Error
                        && double.TryParse(cached, NumberStyles.Float, CultureInfo.InvariantCulture, out double cachedNumber)) {
                        formulaValue = cachedNumber;
                    }
                }
                return new ExcelCellData(ExcelCellDataKind.Formula, formulaValue, ExcelWorksheetScan.FindFormula(cell)!.Text, cached);
            }

            if (ExcelWorksheetScan.FindValue(cell) == null && cell.InlineString == null) {
                return new ExcelCellData(ExcelCellDataKind.Blank, null);
            }

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

        internal void FormatRange(string a1Range, string numberFormat) {
            var (r1, c1, r2, c2) = ExcelA1.ParseRange(a1Range);

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLock(() => FormatRangeCore(r1, c1, r2, c2, numberFormat));
        }

        internal void FillRange(string a1Range, string hexColor) {
            if (string.IsNullOrWhiteSpace(hexColor)) return;
            var (r1, c1, r2, c2) = ExcelA1.ParseRange(a1Range);

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLock(() => FillRangeCore(r1, c1, r2, c2, hexColor));
        }

        internal void ClearRange(string a1Range, ExcelClearOptions options = ExcelClearOptions.All) {
            var (r1, c1, r2, c2) = ExcelA1.ParseRange(a1Range);
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

                if (options.HasFlag(ExcelClearOptions.Values) || options.HasFlag(ExcelClearOptions.Formulas))
                    EnsureDynamicArrayRangeWritable(r1, c1, r2, c2);

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
                    if (options.HasFlag(ExcelClearOptions.Values)
                        || options.HasFlag(ExcelClearOptions.Formulas)) {
                        _excelDocument.MarkFormulaInputMutation();
                    }
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

            foreach ((Cell cell, int rowIndex, int columnIndex) in EnumerateCellsWithEffectiveCoordinates()) {
                if (rowIndex < bounds.r1
                    || rowIndex > bounds.r2
                    || columnIndex < bounds.c1
                    || columnIndex > bounds.c2) {
                    continue;
                }

                if (clearValues && (ExcelWorksheetScan.FindValue(cell) != null
                    || cell.DataType != null
                    || cell.InlineString != null
                    || HasCellValueMetadata(cell))) {
                    cell.CellValue = null;
                    cell.DataType = null;
                    cell.InlineString = null;
                    ClearCellValueMetadata(cell);
                    changed = true;
                }

                if (clearFormulas && ExcelWorksheetScan.FindFormula(cell) != null) {
                    cell.CellFormula = null;
                    changed = true;
                }

                if (clearStyles && cell.StyleIndex != null) {
                    cell.StyleIndex = null;
                    changed = true;
                }
            }

            return changed;
        }
    }
}
