using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryApplyPlainCellsByAppendingRows(IReadOnlyList<(int Row, int Column, object Value)> source, CancellationToken ct) {
            bool applied = false;
            System.Threading.ReaderWriterLockSlim? lck = _excelDocument._lock;
            if (lck == null) {
                try { lck = _excelDocument.EnsureLock(); } catch { lck = null; }
            }

            Locking.ExecuteWrite(lck, () => applied = TryApplyPlainCellsByAppendingRowsCore(source, ct));
            return applied;
        }

        private bool TryApplyPlainCellsByAppendingRowsCore(IReadOnlyList<(int Row, int Column, object Value)> source, CancellationToken ct) {
            if (!TryGetPlainAppendLayout(source, out int firstRow, out int minColumn, out int maxColumn)) {
                return false;
            }

            var sheetData = GetOrCreateSheetData();
            int minExistingRow = int.MaxValue;
            int minExistingColumn = int.MaxValue;
            int maxExistingRow = 0;
            int maxExistingColumn = 0;
            foreach (var existingRow in sheetData.Elements<Row>()) {
                if (existingRow.RowIndex == null) {
                    return false;
                }

                if (existingRow.RowIndex != null && existingRow.RowIndex.Value >= (uint)firstRow) {
                    return false;
                }

                if (!existingRow.HasChildren) {
                    continue;
                }

                int existingRowIndex = checked((int)(existingRow.RowIndex?.Value ?? 0U));
                if (existingRowIndex <= 0) {
                    continue;
                }

                foreach (var existingCell in existingRow.Elements<Cell>()) {
                    int existingColumnIndex = 0;
                    string? reference = existingCell.CellReference?.Value;
                    if (!string.IsNullOrEmpty(reference)) {
                        existingColumnIndex = A1.ParseColumnIndexFromCellReference(reference!);
                    }

                    if (existingColumnIndex <= 0) {
                        continue;
                    }

                    if (existingRowIndex < minExistingRow) minExistingRow = existingRowIndex;
                    if (existingRowIndex > maxExistingRow) maxExistingRow = existingRowIndex;
                    if (existingColumnIndex < minExistingColumn) minExistingColumn = existingColumnIndex;
                    if (existingColumnIndex > maxExistingColumn) maxExistingColumn = existingColumnIndex;
                }
            }

            var columnNames = new string[maxColumn + 1];
            for (int column = 1; column <= maxColumn; column++) {
                columnNames[column] = GetColumnName(column);
            }

            Dictionary<string, int>? sharedStringIndexes = null;
            bool useDirectStringCells = source.Count >= 4096 && maxColumn > 1;
            List<Row> pendingRows = new List<Row>();
            Row? row = null;
            int rowIndex = 0;
            string rowReference = string.Empty;
            bool canCancel = ct.CanBeCanceled;

            for (int i = 0; i < source.Count; i++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var item = source[i];

                if (item.Row != rowIndex) {
                    if (row != null) {
                        pendingRows.Add(row);
                    }

                    rowIndex = item.Row;
                    rowReference = InvariantNumberText.Get(rowIndex);
                    row = new Row { RowIndex = (uint)rowIndex };
                }

                var (cellValue, cellType) = CoercePlainAppendValue(item.Value, ref sharedStringIndexes, useDirectStringCells);
                row!.Append(CreateTabularAppendCell(columnNames[item.Column] + rowReference, cellValue, cellType));
            }

            if (row != null) {
                pendingRows.Add(row);
            }

            foreach (var pendingRow in pendingRows) {
                sheetData.Append(pendingRow);
            }

            ClearHeaderCacheForPreparedAppend();
            int lastRow = source[source.Count - 1].Row;
            int dimensionMinRow = minExistingRow == int.MaxValue ? firstRow : Math.Min(minExistingRow, firstRow);
            int dimensionMinColumn = minExistingColumn == int.MaxValue ? minColumn : Math.Min(minExistingColumn, minColumn);
            int dimensionMaxRow = Math.Max(maxExistingRow, lastRow);
            int dimensionMaxColumn = Math.Max(maxExistingColumn, maxColumn);
            SetSheetDimensionReference(dimensionMinRow, dimensionMinColumn, dimensionMaxRow, dimensionMaxColumn);
            _requiresSavePreparation = false;
            return true;
        }

        private void ClearHeaderCacheForPreparedAppend() {
            _hasWorksheetMutations = true;
            _excelDocument.MarkPackageDirty();
            ClearCellTextSharedStringCache();
            lock (_headerMapLock) {
                _headerMapCache = null;
                _headerMapSourceA1 = null;
            }
        }

        private bool TryGetPlainAppendLayout(
            IReadOnlyList<(int Row, int Column, object Value)> source,
            out int firstRow,
            out int minColumn,
            out int maxColumn) {
            firstRow = source[0].Row;
            minColumn = int.MaxValue;
            maxColumn = 0;
            int currentRow = 0;
            int currentColumn = 0;

            for (int i = 0; i < source.Count; i++) {
                var item = source[i];
                if (item.Row <= 0 || item.Column <= 0 || item.Column > A1.MaxColumns || item.Row < currentRow) {
                    return false;
                }

                if (!CanAppendPlainValueDirectly(item.Value)) {
                    return false;
                }

                if (item.Row != currentRow) {
                    currentRow = item.Row;
                    currentColumn = 0;
                }

                if (item.Column <= currentColumn) {
                    return false;
                }

                currentColumn = item.Column;
                if (item.Column < minColumn) {
                    minColumn = item.Column;
                }

                if (item.Column > maxColumn) {
                    maxColumn = item.Column;
                }
            }

            if (minColumn == int.MaxValue) {
                minColumn = 1;
            }

            return true;
        }

        private void SetSheetDimensionReference(int minRow, int minColumn, int maxRow, int maxColumn) {
            var worksheet = WorksheetRoot;
            SheetDimension? dimension = null;
            List<SheetDimension>? extraDimensions = null;
            foreach (var currentDimension in worksheet.Elements<SheetDimension>()) {
                if (dimension == null) {
                    dimension = currentDimension;
                    continue;
                }

                extraDimensions ??= new List<SheetDimension>();
                extraDimensions.Add(currentDimension);
            }

            if (extraDimensions != null) {
                for (int i = 0; i < extraDimensions.Count; i++) {
                    extraDimensions[i].Remove();
                }
            }

            string start = A1.CellReference(minRow, minColumn);
            string end = A1.CellReference(maxRow, maxColumn);
            string reference = start == end ? start : start + ":" + end;
            if (dimension == null) {
                InsertSheetDimensionInSchemaOrder(worksheet, new SheetDimension { Reference = reference });
            } else {
                dimension.Reference = reference;
            }
        }

        private static bool CanAppendPlainValueDirectly(object? value) {
            switch (value) {
                case null:
                case DBNull:
                case double:
                case float:
                case decimal:
                case int:
                case long:
                case bool:
                case uint:
                case ulong:
                case ushort:
                case byte:
                case sbyte:
                case short:
                case Guid:
                case Enum:
                case char:
                case Uri:
                    return true;
                case string text:
                    if (text.IndexOf('\r') >= 0 || text.IndexOf('\n') >= 0) {
                        return false;
                    }

                    CoerceValueHelper.ValidateSharedStringLength(text, nameof(value));
                    return true;
                default:
                    return false;
            }
        }

        private (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) CoercePlainAppendValue(
            object? value,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool useDirectStringCells) {
            (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) = value switch {
                null => CoerceValueHelper.HandleEmptyString(),
                DBNull => CoerceValueHelper.HandleEmptyString(),
                string text when text.Length == 0 => CoerceValueHelper.HandleEmptyString(),
                string text => useDirectStringCells
                    ? (CreatePrevalidatedPlainAppendStringValue(text), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                    : (CreatePlainAppendSharedStringValue(text, ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString),
                double number => CoerceValueHelper.HandleNumber(number),
                float number => CoerceValueHelper.HandleNumber((double)number),
                decimal number => CoerceValueHelper.HandleDecimal(number),
                int number => CoerceValueHelper.HandleSignedInteger(number),
                long number => CoerceValueHelper.HandleSignedInteger(number),
                bool flag => CoerceValueHelper.HandleBoolean(flag),
                uint number => CoerceValueHelper.HandleUnsignedInteger(number),
                ulong number => CoerceValueHelper.HandleUnsignedInteger(number),
                ushort number => CoerceValueHelper.HandleUnsignedInteger(number),
                byte number => CoerceValueHelper.HandleUnsignedInteger(number),
                sbyte number => CoerceValueHelper.HandleSignedInteger(number),
                short number => CoerceValueHelper.HandleSignedInteger(number),
                Guid guid => useDirectStringCells
                    ? (CreatePlainAppendStringValue(guid.ToString()), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                    : (CreatePlainAppendSharedStringValue(guid.ToString(), ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString),
                Enum enumValue => useDirectStringCells
                    ? (CreatePlainAppendStringValue(enumValue.ToString()), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                    : (CreatePlainAppendSharedStringValue(enumValue.ToString(), ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString),
                char character => useDirectStringCells
                    ? (CreatePlainAppendStringValue(character.ToString()), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                    : (CreatePlainAppendSharedStringValue(character.ToString(), ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString),
                Uri uri => useDirectStringCells
                    ? (CreatePlainAppendStringValue(uri.ToString()), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                    : (CreatePlainAppendSharedStringValue(uri.ToString(), ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString),
                _ => throw new InvalidOperationException("Unsupported direct append value.")
            };

            return (cellValue, cellType);
        }

        private static CellValue CreatePlainAppendStringValue(string text) {
            CoerceValueHelper.ValidateSharedStringLength(text, nameof(text));
            return new CellValue(Utilities.ExcelSanitizer.SanitizeString(text));
        }

        private static CellValue CreatePrevalidatedPlainAppendStringValue(string text) {
            return new CellValue(Utilities.ExcelSanitizer.SanitizeString(text));
        }

        private CellValue CreatePlainAppendSharedStringValue(string text, ref Dictionary<string, int>? sharedStringIndexes) {
            string sanitized = Utilities.ExcelSanitizer.SanitizeString(text);
            sharedStringIndexes ??= new Dictionary<string, int>(StringComparer.Ordinal);
            if (!sharedStringIndexes.TryGetValue(sanitized, out int index)) {
                index = _excelDocument.GetSharedStringIndex(sanitized);
                sharedStringIndexes[sanitized] = index;
            }

            return new CellValue(SharedStringIndexText.Get(index));
        }
    }
}
