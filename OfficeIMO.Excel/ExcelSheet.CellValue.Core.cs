using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        private readonly struct CellValueSharedStringIndexCacheEntry {
            internal CellValueSharedStringIndexCacheEntry(int index, bool containsLineBreak) {
                Index = index;
                ContainsLineBreak = containsLineBreak;
            }

            internal int Index { get; }

            internal bool ContainsLineBreak { get; }
        }

        private void CellValueCoreNoMaterialize(int row, int column, object? value) {
            if (value == null || value == DBNull.Value) {
                CellEmptyStringValueCore(row, column);
                return;
            }

            switch (value) {
                case string text:
                    CellStringValueCore(row, column, text);
                    return;
                case double number:
                    CellDoubleValueCore(row, column, number);
                    return;
                case float number:
                    CellDoubleValueCore(row, column, (double)number);
                    return;
                case decimal number:
                    CellDecimalValueCore(row, column, number);
                    return;
                case int number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case long number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case short number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case uint number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case ulong number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case ushort number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case byte number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case sbyte number:
                    CellNumberTextValueCore(row, column, InvariantNumberText.Get(number));
                    return;
                case bool boolean:
                    CellBooleanValueCore(row, column, boolean);
                    return;
                case DateTime dateTime:
                    CellDateTimeValueCore(row, column, dateTime);
                    return;
                case DateTimeOffset dateTimeOffset:
                    CellDateTimeOffsetValueCore(row, column, dateTimeOffset);
                    return;
#if NET6_0_OR_GREATER
                case DateOnly dateOnly:
                    CellDateOnlyValueCore(row, column, dateOnly);
                    return;
                case TimeOnly timeOnly:
                    CellTimeOnlyValueCore(row, column, timeOnly);
                    return;
#endif
                case TimeSpan timeSpan:
                    CellTimeSpanValueCore(row, column, timeSpan);
                    return;
            }

            var (cellValue, dataType) = CoerceForCell(value);

            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            cell.CellValue = cellValue;
            cell.DataType = dataType;
            ApplyAutomaticCellFormatting(cell, value, dataType);
            CompleteCellValueMutation(row, column);
        }

        private void CellStringValueCore(int row, int column, string? value) {
            if (string.IsNullOrEmpty(value)) {
                CellEmptyStringValueCore(row, column);
                return;
            }

            CoerceValueHelper.ValidateSharedStringLength(value!, nameof(value));
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            string text = value!;
            if (TryGetCellValueSharedStringIndex(text, out int cachedSharedStringIndex, out bool cachedContainsLineBreak)) {
                SetExistingCellSharedStringValue(cell, cachedSharedStringIndex, cachedContainsLineBreak);
            } else if (_excelDocument.TryGetOrAddSharedStringIndexBelowLimit(
                    text,
                    CellValuePlainStringPromotionSharedStringCount,
                    validateNewString: true,
                    out int sharedStringIndex,
                    out bool containsLineBreak)) {
                AddCellValueSharedStringIndex(text, sharedStringIndex, containsLineBreak);
                SetExistingCellSharedStringValue(cell, sharedStringIndex, containsLineBreak);
            } else {
                SetExistingCellPlainStringValue(cell, text);
            }

            CompleteCellValueMutation(row, column);
        }

        private bool TryGetCellValueSharedStringIndex(string text, out int index, out bool containsLineBreak) {
            SharedStringTable? table = _excelDocument.WorkbookPartRoot.SharedStringTablePart?.SharedStringTable;
            if (_cellValueSharedStringIndexCache != null && (_cellValueSharedStringTable == null
                || !_cellValueSharedStringTable.TryGetTarget(out SharedStringTable? cachedTable)
                || !ReferenceEquals(cachedTable, table))) {
                _cellValueSharedStringIndexCache = null;
            }
            if (_cellValueSharedStringIndexCache != null
                && _cellValueSharedStringIndexCache.TryGetValue(text, out CellValueSharedStringIndexCacheEntry entry)) {
                index = entry.Index;
                containsLineBreak = entry.ContainsLineBreak;
                return true;
            }

            index = -1;
            containsLineBreak = false;
            return false;
        }

        private void AddCellValueSharedStringIndex(string text, int index, bool containsLineBreak) {
            SharedStringTable? table = _excelDocument.WorkbookPartRoot.SharedStringTablePart?.SharedStringTable;
            if (table == null) return;
            if (_cellValueSharedStringTable == null) _cellValueSharedStringTable = new WeakReference<SharedStringTable>(table);
            else _cellValueSharedStringTable.SetTarget(table);
            var cache = _cellValueSharedStringIndexCache;
            if (cache == null) {
                cache = new Dictionary<string, CellValueSharedStringIndexCacheEntry>(StringComparer.Ordinal);
                _cellValueSharedStringIndexCache = cache;
            } else if (cache.Count >= CellValueSharedStringIndexCacheLimit) {
                return;
            }

            cache[text] = new CellValueSharedStringIndexCacheEntry(index, containsLineBreak);
        }

        private void CellEmptyStringValueCore(int row, int column) {
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            cell.CellValue = new CellValue(string.Empty);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            cell.InlineString = null;
            CompleteCellValueMutation(row, column);
        }

        private void SetExistingCellSharedStringValue(Cell cell, string value, int sharedStringIndex) {
            SetExistingCellSharedStringValue(cell, sharedStringIndex, value.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0);
        }

        private void SetExistingCellSharedStringValue(Cell cell, int sharedStringIndex, bool containsLineBreak) {
            cell.CellValue = new CellValue(SharedStringIndexText.Get(sharedStringIndex));
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
            cell.InlineString = null;
            if (containsLineBreak) {
                ApplyWrapText(cell);
            }
        }

        private void SetExistingCellPlainStringValue(Cell cell, string value) {
            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            cell.CellValue = new CellValue(value);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            cell.InlineString = null;
            if (value.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0) {
                ApplyWrapText(cell);
            }
        }

        private void CellDoubleValueCore(int row, int column, double value) {
            string text = InvariantNumberText.Get(value);
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            cell.CellValue = new CellValue(text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            CompleteCellValueMutation(row, column);
        }

        private void CellDecimalValueCore(int row, int column, decimal value) {
            string text = value.ToString(CultureInfo.InvariantCulture);
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            cell.CellValue = new CellValue(text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            CompleteCellValueMutation(row, column);
        }

        private void CellNumberTextValueCore(int row, int column, string text) {
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            cell.CellValue = new CellValue(text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            CompleteCellValueMutation(row, column);
        }

        private void CellBooleanValueCore(int row, int column, bool value) {
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            cell.CellValue = new CellValue(value ? "1" : "0");
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean;
            CompleteCellValueMutation(row, column);
        }

        private void CellDateTimeValueCore(int row, int column, DateTime value) {
            double serial = ExcelDateSystemConverter.ToSerial(value, _excelDocument.DateSystem);
            string text = InvariantNumberText.Get(serial);
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
            cell.CellValue = new CellValue(text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            cell.StyleIndex = GetAutomaticNumberFormatStyleIndex(baseStyleIndex, duration: false);
            CompleteCellValueMutation(row, column);
        }

        private DateTime ConvertDateTimeOffsetForCell(DateTimeOffset value) {
            try {
                return _excelDocument.DateTimeOffsetWriteStrategy(value);
            } catch (Exception ex) {
                throw new InvalidOperationException("The configured DateTimeOffset write strategy threw an exception.", ex);
            }
        }

        private void CellDateTimeOffsetValueCore(int row, int column, DateTimeOffset value) =>
            CellDateTimeOffsetValueCore(row, column, value, ConvertDateTimeOffsetForCell(value));

        private void CellDateTimeOffsetValueCore(int row, int column, DateTimeOffset value, DateTime converted) {
            string? numericText = null;
            if (value.UtcDateTime >= CellValueExcelMinimumSupportedDate) {
                try {
                    double serial = ExcelDateSystemConverter.ToSerial(converted, _excelDocument.DateSystem);
                    numericText = InvariantNumberText.Get(serial);
                } catch (ArgumentException) {
                    // Fall back to ISO text below for values Excel cannot represent numerically.
                } catch (OverflowException) {
                    // Fall back to ISO text below for values Excel cannot represent numerically.
                }
            }

            if (numericText != null) {
                var numericCell = GetWritableValueCell(row, column);
                ClearCellValueMetadata(numericCell);
                numericCell.CellValue = new CellValue(numericText);
                numericCell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;

                uint baseStyleIndex = numericCell.StyleIndex?.Value ?? 0U;
                numericCell.StyleIndex = GetAutomaticNumberFormatStyleIndex(baseStyleIndex, duration: false);

                CompleteCellValueMutation(row, column);
                return;
            }

            string fallbackText = value.ToString("o", CultureInfo.InvariantCulture);
            int sharedStringIndex = _excelDocument.GetSharedStringIndex(fallbackText, validateNewString: true, out bool containsLineBreak);
            var fallbackCell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(fallbackCell);
            SetExistingCellSharedStringValue(fallbackCell, sharedStringIndex, containsLineBreak);
            CompleteCellValueMutation(row, column);
        }

#if NET6_0_OR_GREATER
        private void CellDateOnlyValueCore(int row, int column, DateOnly value) {
            string text = InvariantNumberText.Get(ExcelDateSystemConverter.ToSerial(
                value.ToDateTime(TimeOnly.MinValue),
                _excelDocument.DateSystem));
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
            cell.CellValue = new CellValue(text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            cell.StyleIndex = GetAutomaticNumberFormatStyleIndex(baseStyleIndex, duration: false);
            CompleteCellValueMutation(row, column);
        }

        private void CellTimeOnlyValueCore(int row, int column, TimeOnly value) {
            string text = InvariantNumberText.Get(value.ToTimeSpan().TotalDays);
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
            cell.CellValue = new CellValue(text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            cell.StyleIndex = GetAutomaticNumberFormatStyleIndex(baseStyleIndex, duration: true);
            CompleteCellValueMutation(row, column);
        }
#endif

        private void CellFormulaCore(int row, int column, string formula) {
            var safe = Utilities.ExcelSanitizer.SanitizeFormula(formula);
            Cell cell = GetWritableValueCell(row, column);
            bool retainsCachedValue = cell.CellValue != null;
            ClearCellValueMetadata(cell);
            // Excel formulas in XML should not start with '=' and must not include illegal control characters
            cell.CellFormula = new CellFormula(safe);
            if (retainsCachedValue) {
                cell.CellFormula.CalculateCell = true;
            }
            _excelDocument.MarkFormulaAuthored(
                _worksheetPart,
                A1.CellReference(row, column),
                retainedCachedValue: retainsCachedValue);
            ClearHeaderCacheForCellMutation(row, column);
        }

        private void CellTimeSpanValueCore(int row, int column, TimeSpan value) {
            double serial = value.TotalDays;
            string text = InvariantNumberText.Get(serial);
            var cell = GetWritableValueCell(row, column);
            ClearCellValueMetadata(cell);
            uint baseStyleIndex = cell.StyleIndex?.Value ?? 0U;
            cell.CellValue = new CellValue(text);
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
            cell.StyleIndex = GetAutomaticNumberFormatStyleIndex(baseStyleIndex, duration: true);
            CompleteCellValueMutation(row, column);
        }

        private void CompleteCellValueMutation(int row, int column) {
            _excelDocument.MarkFormulaInputMutation();
            ClearHeaderCacheForCellMutation(row, column);
        }

        // Core coercion logic shared between sequential and parallel operations
        private (CellValue cellValue, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> dataType) CoerceForCell(object? value) {
            var dateTimeOffsetStrategy = _excelDocument.DateTimeOffsetWriteStrategy;
            var (cellValue, cellType) = CoerceValueHelper.Coerce(
                value,
                s => {
                    int idx = _excelDocument.GetSharedStringIndex(s);
                    return new CellValue(SharedStringIndexText.Get(idx));
                },
                dateTimeOffsetStrategy,
                _excelDocument.DateSystem);
            return (cellValue, new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cellType));
        }
    }
}
