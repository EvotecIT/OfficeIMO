using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool UsesCellStore => Persistence.UsesCellStore;

        /// <summary>
        /// Finds a cell for reading without moving a stored cell into the DOM. Returns the DOM cell, or null with
        /// <paramref name="storedRow"/> set when the cell is stored, or null with no row when the cell does not exist.
        /// </summary>
        private Cell? ProbeCell(int row, int column, out ExcelStoredRow? storedRow, out int storedIndex) {
            storedRow = null;
            storedIndex = -1;
            Row? rowElement = TryGetRowElement(row, column);
            if (rowElement == null) return null;
            if (rowElement is ExcelStoredRow stored && stored.IsStoreAvailable) {
                int index = stored.Find(column);
                if (index >= 0) {
                    storedRow = stored;
                    storedIndex = index;
                    return null;
                }
                using (ExcelStoredRow.BeginRawAccess()) {
                    return TryGetCellInRow(rowElement, row, column);
                }
            }
            return TryGetCellInRow(rowElement, row, column);
        }

        /// <summary>Reads a stored cell as <see cref="GetCellValueSnapshot(Cell?)"/> reads its DOM form.</summary>
        private ExcelCellData GetStoredCellValueSnapshot(in ExcelStoredCell cell) {
            if (!cell.HasValue) return new ExcelCellData(ExcelCellDataKind.Blank, null);
            string cached = cell.Value!;
            switch (cell.Type) {
                case ExcelStoredCellType.Boolean:
                    return new ExcelCellData(ExcelCellDataKind.Boolean, cached == "1" || string.Equals(cached, "true", StringComparison.OrdinalIgnoreCase), cachedText: cached);
                case ExcelStoredCellType.Error:
                    return new ExcelCellData(ExcelCellDataKind.Error, cached, cachedText: cached);
                case ExcelStoredCellType.SharedString:
                    string shared = TryParseCellTextSharedStringIndex(cached, out int index)
                        ? BuildCellTextSharedStringSnapshot().Get(index) ?? string.Empty
                        : string.Empty;
                    return new ExcelCellData(ExcelCellDataKind.Text, shared, cachedText: shared);
                case ExcelStoredCellType.String:
                    return new ExcelCellData(ExcelCellDataKind.Text, cached, cachedText: cached);
            }

            if (double.TryParse(cached, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                return new ExcelCellData(ExcelCellDataKind.Number, number, cachedText: cached);
            }

            return cached.Length == 0
                ? new ExcelCellData(ExcelCellDataKind.Blank, null)
                : new ExcelCellData(ExcelCellDataKind.Text, cached, cachedText: cached);
        }

        /// <summary>Reads a stored number or shared-text cell as <see cref="TryResolvePlainNumberCell"/> reads its DOM form.</summary>
        private bool TryResolveStoredPlainNumberCell(in ExcelStoredCell cell, out FormulaArgumentValue value) {
            value = default;
            string? text = cell.Value;
            if (cell.Type == ExcelStoredCellType.Absent || cell.Type == ExcelStoredCellType.Number) {
                if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) return false;
                value = new FormulaArgumentValue(number, text, sourceCellKind: ExcelCellDataKind.Number);
                return true;
            }
            if (cell.Type != ExcelStoredCellType.SharedString
                || string.IsNullOrEmpty(text) || !TryParseCellTextSharedStringIndex(text!, out int index)) return false;
            value = new FormulaArgumentValue(null, BuildCellTextSharedStringSnapshot().Get(index) ?? string.Empty,
                sourceCellKind: ExcelCellDataKind.Text);
            return true;
        }

        /// <summary>
        /// Finds where a plain value for a cell is kept: an existing stored cell, or the place for a new one in a
        /// stored row. Returns false when the cell is a DOM cell or the worksheet does not store cells, and the value
        /// goes through the DOM.
        /// </summary>
        private bool TryGetStoredCellForWrite(int row, int column, out ExcelStoredRow storedRow, out int index) {
            storedRow = null!;
            index = -1;
            if (!UsesCellStore) return false;

            EnsureDynamicArrayCellWritable(row, column);
            SheetData sheetData = GetOrCreateSheetData();
            Row rowElement = GetOrCreateRowForCell(sheetData, row, out _);
            if (rowElement is not ExcelStoredRow stored || !stored.AcceptsStoredCells) return false;

            int found = stored.Find(column);
            if (found < 0) {
                using (ExcelStoredRow.BeginRawAccess()) {
                    if (TryGetCellInRow(rowElement, row, column) != null) return false;
                }
            }

            storedRow = stored;
            index = found;
            return true;
        }

        /// <summary>Gets the style a stored cell keeps when its value changes, or a new cell inherits from its row or column.</summary>
        private uint GetStoredWriteBaseStyle(ExcelStoredRow storedRow, int index, int column, out bool hasStyle) {
            if (index >= 0) {
                ref ExcelStoredCell existing = ref storedRow.CellAt(index);
                hasStyle = existing.HasStyle;
                return existing.Style;
            }

            Cell scratch = StyleScratchCell(0U);
            ApplyInheritedCellStyle(scratch, GetOrCreateSheetData(), storedRow, column);
            hasStyle = scratch.StyleIndex != null;
            return scratch.StyleIndex?.Value ?? 0U;
        }

        private void SetStoredValue(ExcelStoredRow storedRow, int index, int column, uint style, bool hasStyle, ExcelStoredCellType type, string value) {
            var cell = new ExcelStoredCell(column, style, hasStyle, type, value);
            if (index >= 0) storedRow.CellAt(index) = cell;
            else storedRow.Insert(~index, cell);
        }

        /// <summary>Writes a number given as invariant text, as the DOM route's numeric writers do.</summary>
        private bool TryStoreNumberText(int row, int column, string text) {
            if (!TryGetStoredCellForWrite(row, column, out ExcelStoredRow storedRow, out int index)) return false;
            uint style = GetStoredWriteBaseStyle(storedRow, index, column, out bool hasStyle);
            SetStoredValue(storedRow, index, column, style, hasStyle, ExcelStoredCellType.Number, text);
            CompleteCellValueMutation(row, column);
            return true;
        }

        /// <summary>Writes a date or time serial with the automatic number format, as the DOM route does.</summary>
        private bool TryStoreTemporal(int row, int column, string text, bool duration) {
            if (!TryGetStoredCellForWrite(row, column, out ExcelStoredRow storedRow, out int index)) return false;
            uint baseStyle = GetStoredWriteBaseStyle(storedRow, index, column, out _);
            uint style = GetAutomaticNumberFormatStyleIndex(baseStyle, duration);
            SetStoredValue(storedRow, index, column, style, hasStyle: true, ExcelStoredCellType.Number, text);
            CompleteCellValueMutation(row, column);
            return true;
        }

        private bool TryStoreBoolean(int row, int column, bool value) {
            if (!TryGetStoredCellForWrite(row, column, out ExcelStoredRow storedRow, out int index)) return false;
            uint style = GetStoredWriteBaseStyle(storedRow, index, column, out bool hasStyle);
            SetStoredValue(storedRow, index, column, style, hasStyle, ExcelStoredCellType.Boolean, value ? "1" : "0");
            CompleteCellValueMutation(row, column);
            return true;
        }

        /// <summary>
        /// Writes non-empty text as a shared string, or as cell text past the shared-string promotion limit, with the
        /// wrap format a line break needs, as <see cref="CellStringValueCore"/> does.
        /// </summary>
        private bool TryStoreText(int row, int column, string text) {
            if (!TryGetStoredCellForWrite(row, column, out ExcelStoredRow storedRow, out int index)) return false;
            uint style = GetStoredWriteBaseStyle(storedRow, index, column, out bool hasStyle);
            ExcelStoredCellType type;
            string value;
            bool containsLineBreak;
            if (TryGetCellValueSharedStringIndex(text, out int cachedIndex, out bool cachedLineBreak)) {
                type = ExcelStoredCellType.SharedString;
                value = SharedStringIndexText.Get(cachedIndex);
                containsLineBreak = cachedLineBreak;
            } else if (_excelDocument.TryGetOrAddSharedStringIndexBelowLimit(
                    text,
                    CellValuePlainStringPromotionSharedStringCount,
                    validateNewString: true,
                    out int sharedStringIndex,
                    out bool sharedLineBreak)) {
                AddCellValueSharedStringIndex(text, sharedStringIndex, sharedLineBreak);
                type = ExcelStoredCellType.SharedString;
                value = SharedStringIndexText.Get(sharedStringIndex);
                containsLineBreak = sharedLineBreak;
            } else {
                type = ExcelStoredCellType.String;
                value = text;
                containsLineBreak = text.IndexOf('\n') >= 0 || text.IndexOf('\r') >= 0;
            }

            if (containsLineBreak) {
                Cell scratch = StyleScratchCell(style);
                if (!hasStyle) scratch.StyleIndex = null;
                ApplyWrapText(scratch);
                hasStyle = scratch.StyleIndex != null;
                style = scratch.StyleIndex?.Value ?? 0U;
            }

            SetStoredValue(storedRow, index, column, style, hasStyle, type, value);
            CompleteCellValueMutation(row, column);
            return true;
        }
    }
}
