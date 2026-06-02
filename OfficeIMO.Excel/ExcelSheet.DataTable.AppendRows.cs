using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryInsertDataTableByAppendingRows(DataTable table, int startRow, int startColumn, bool includeHeaders, CancellationToken ct) {
            int columnCount = table.Columns.Count;
            int rowsCount = table.Rows.Count + (includeHeaders ? 1 : 0);
            if (columnCount == 0 || rowsCount == 0) {
                return true;
            }

            if (startColumn + columnCount - 1 > A1.MaxColumns || startRow + rowsCount - 1 > A1.MaxRows) {
                return false;
            }

            bool applied = false;
            System.Threading.ReaderWriterLockSlim? lck = _excelDocument._lock;
            if (lck == null) {
                try { lck = _excelDocument.EnsureLock(); } catch { lck = null; }
            }

            Locking.ExecuteWrite(lck, () => applied = TryInsertDataTableByAppendingRowsCore(table, startRow, startColumn, includeHeaders, ct));
            return applied;
        }

        private bool TryInsertDataTableByAppendingRowsCore(DataTable table, int startRow, int startColumn, bool includeHeaders, CancellationToken ct) {
            var sheetData = GetOrCreateSheetData();
            int minExistingRow = int.MaxValue;
            int minExistingColumn = int.MaxValue;
            int maxExistingRow = 0;
            int maxExistingColumn = 0;
            foreach (var existingRow in sheetData.Elements<Row>()) {
                if (existingRow.RowIndex == null) {
                    return false;
                }

                int existingRowIndex = checked((int)(existingRow.RowIndex?.Value ?? 0U));
                if (existingRowIndex >= startRow) {
                    return false;
                }

                if (existingRowIndex <= 0 || !existingRow.HasChildren) {
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

            int columnCount = table.Columns.Count;
            string[] columnReferencePrefixes = BuildColumnReferencePrefixes(startColumn, columnCount);

            string?[] numberFormats = BuildDataTableNumberFormats(table);
            var stylePlanner = new StylePlanner();
            bool hasObjectColumn = false;
            foreach (DataColumn column in table.Columns) {
                if (column.DataType == typeof(object)) {
                    hasObjectColumn = true;
                    break;
                }
            }

            foreach (string? numberFormat in numberFormats) {
                stylePlanner.NoteNumberFormat(numberFormat);
            }

            if (hasObjectColumn) {
                stylePlanner.NoteNumberFormat(DataTableDateTimeNumberFormat);
                stylePlanner.NoteNumberFormat(DataTableTimeSpanNumberFormat);
            }

            stylePlanner.ApplyTo(_excelDocument);
            var styleIndexes = new uint?[numberFormats.Length];
            for (int i = 0; i < numberFormats.Length; i++) {
                if (stylePlanner.TryGetCellFormatIndex(numberFormats[i], out uint styleIndex)) {
                    styleIndexes[i] = styleIndex;
                }
            }

            uint? objectDateTimeStyleIndex = null;
            uint? objectTimeSpanStyleIndex = null;
            if (hasObjectColumn) {
                if (stylePlanner.TryGetCellFormatIndex(DataTableDateTimeNumberFormat, out uint dateTimeStyleIndex)) {
                    objectDateTimeStyleIndex = dateTimeStyleIndex;
                }

                if (stylePlanner.TryGetCellFormatIndex(DataTableTimeSpanNumberFormat, out uint timeSpanStyleIndex)) {
                    objectTimeSpanStyleIndex = timeSpanStyleIndex;
                }
            }

            var columnKinds = BuildDataTableAppendColumnKinds(table);
            int cellCount = (table.Rows.Count + (includeHeaders ? 1 : 0)) * columnCount;
            bool useDirectStringCells = false;
            Dictionary<string, int>? sharedStringIndexes = null;
            int rowIndex = startRow;
            bool canCancel = ct.CanBeCanceled;
            List<Row>? pendingRows = canCancel ? new List<Row>(table.Rows.Count + (includeHeaders ? 1 : 0)) : null;

            if (includeHeaders) {
                Row headerRow = CreateDataTableHeaderRow(rowIndex++, columnReferencePrefixes, table, useDirectStringCells, ref sharedStringIndexes, canCancel, ct);
                if (pendingRows != null) {
                    pendingRows.Add(headerRow);
                } else {
                    sheetData.Append(headerRow);
                }
            }

            foreach (DataRow dataRow in table.Rows) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                Row valueRow = canCancel
                    ? CreateDataTableValueRow(rowIndex++, columnReferencePrefixes, dataRow, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes, canCancel, ct)
                    : CreateDataTableValueRow(rowIndex++, columnReferencePrefixes, dataRow, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes);
                if (pendingRows != null) {
                    pendingRows.Add(valueRow);
                } else {
                    sheetData.Append(valueRow);
                }
            }

            if (pendingRows != null) {
                foreach (var pendingRow in pendingRows) {
                    sheetData.Append(pendingRow);
                }
            }

            ClearHeaderCacheForPreparedAppend();
            int lastRow = startRow + table.Rows.Count + (includeHeaders ? 1 : 0) - 1;
            int lastColumn = startColumn + columnCount - 1;
            int dimensionMinRow = minExistingRow == int.MaxValue ? startRow : Math.Min(minExistingRow, startRow);
            int dimensionMinColumn = minExistingColumn == int.MaxValue ? startColumn : Math.Min(minExistingColumn, startColumn);
            int dimensionMaxRow = Math.Max(maxExistingRow, lastRow);
            int dimensionMaxColumn = Math.Max(maxExistingColumn, lastColumn);
            SetSheetDimensionReference(dimensionMinRow, dimensionMinColumn, dimensionMaxRow, dimensionMaxColumn);
            _requiresSavePreparation = false;
            return true;
        }

        internal bool TryInsertTabularRowSourceForDeferredMaterialization(IExcelSheetTabularRowSource source, int startRow = 1, int startColumn = 1, bool includeHeaders = true, CancellationToken ct = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (startRow < 1) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn < 1) throw new ArgumentOutOfRangeException(nameof(startColumn));

            int columnCount = source.ColumnCount;
            int rowsCount = source.RowCount + (includeHeaders ? 1 : 0);
            if (columnCount == 0 || rowsCount == 0) {
                return true;
            }

            if (startColumn + columnCount - 1 > A1.MaxColumns || startRow + rowsCount - 1 > A1.MaxRows) {
                return false;
            }

            bool applied = false;
            System.Threading.ReaderWriterLockSlim? lck = _excelDocument._lock;
            if (lck == null) {
                try { lck = _excelDocument.EnsureLock(); } catch { lck = null; }
            }

            Locking.ExecuteWrite(lck, () => applied = TryInsertTabularRowSourceByAppendingRowsCore(source, startRow, startColumn, includeHeaders, ct));
            return applied;
        }

        private bool TryInsertTabularRowSourceByAppendingRowsCore(IExcelSheetTabularRowSource source, int startRow, int startColumn, bool includeHeaders, CancellationToken ct) {
            var sheetData = GetOrCreateSheetData();
            int minExistingRow = int.MaxValue;
            int minExistingColumn = int.MaxValue;
            int maxExistingRow = 0;
            int maxExistingColumn = 0;
            foreach (var existingRow in sheetData.Elements<Row>()) {
                if (existingRow.RowIndex == null) {
                    return false;
                }

                int existingRowIndex = checked((int)(existingRow.RowIndex?.Value ?? 0U));
                if (existingRowIndex >= startRow) {
                    return false;
                }

                if (existingRowIndex <= 0 || !existingRow.HasChildren) {
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

            int columnCount = source.ColumnCount;
            string[] columnReferencePrefixes = BuildColumnReferencePrefixes(startColumn, columnCount);

            string?[] numberFormats = BuildTabularRowSourceNumberFormats(source);
            var stylePlanner = new StylePlanner();
            bool hasObjectColumn = false;
            for (int i = 0; i < columnCount; i++) {
                if (source.GetColumnType(i) == typeof(object)) {
                    hasObjectColumn = true;
                    break;
                }
            }

            foreach (string? numberFormat in numberFormats) {
                stylePlanner.NoteNumberFormat(numberFormat);
            }

            if (hasObjectColumn) {
                stylePlanner.NoteNumberFormat(DataTableDateTimeNumberFormat);
                stylePlanner.NoteNumberFormat(DataTableTimeSpanNumberFormat);
            }

            stylePlanner.ApplyTo(_excelDocument);
            var styleIndexes = new uint?[numberFormats.Length];
            for (int i = 0; i < numberFormats.Length; i++) {
                if (stylePlanner.TryGetCellFormatIndex(numberFormats[i], out uint styleIndex)) {
                    styleIndexes[i] = styleIndex;
                }
            }

            uint? objectDateTimeStyleIndex = null;
            uint? objectTimeSpanStyleIndex = null;
            if (hasObjectColumn) {
                if (stylePlanner.TryGetCellFormatIndex(DataTableDateTimeNumberFormat, out uint dateTimeStyleIndex)) {
                    objectDateTimeStyleIndex = dateTimeStyleIndex;
                }

                if (stylePlanner.TryGetCellFormatIndex(DataTableTimeSpanNumberFormat, out uint timeSpanStyleIndex)) {
                    objectTimeSpanStyleIndex = timeSpanStyleIndex;
                }
            }

            var columnKinds = BuildTabularAppendColumnKinds(source);
            int rowCount = source.RowCount;
            int cellCount = (rowCount + (includeHeaders ? 1 : 0)) * columnCount;
            bool useDirectStringCells = cellCount >= 4096 && columnCount > 1;
            Dictionary<string, int>? sharedStringIndexes = null;
            int rowIndex = startRow;
            bool canCancel = ct.CanBeCanceled;
            List<Row> pendingRows = new List<Row>(rowCount + (includeHeaders ? 1 : 0));
            object?[]? flatValues = null;
            bool useFlatValues = source.TryGetFlatValues(out var sourceFlatValues, out int flatColumnCount) && flatColumnCount == columnCount;
            if (useFlatValues) {
                flatValues = sourceFlatValues;
            }

            if (includeHeaders) {
                Row headerRow = CreateTabularRowSourceHeaderRow(rowIndex++, columnReferencePrefixes, source, useDirectStringCells, ref sharedStringIndexes, canCancel, ct);
                pendingRows.Add(headerRow);
            }

            for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                Row valueRow;
                if (flatValues != null && !canCancel) {
                    valueRow = CreateTabularRowSourceValueRow(rowIndex++, columnReferencePrefixes, flatValues, sourceRowIndex * columnCount, columnCount, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes);
                } else if (flatValues != null) {
                    valueRow = CreateTabularRowSourceValueRow(rowIndex++, columnReferencePrefixes, flatValues, sourceRowIndex * columnCount, columnCount, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes, canCancel, ct);
                } else if (source.TryGetBufferedRow(sourceRowIndex, out var rowValues) && rowValues != null && !canCancel) {
                    valueRow = CreateTabularRowSourceValueRow(rowIndex++, columnReferencePrefixes, rowValues, 0, columnCount, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes);
                } else if (rowValues != null) {
                    valueRow = CreateTabularRowSourceValueRow(rowIndex++, columnReferencePrefixes, rowValues, 0, columnCount, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes, canCancel, ct);
                } else {
                    valueRow = CreateTabularRowSourceValueRow(rowIndex++, columnReferencePrefixes, source, sourceRowIndex, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes, canCancel, ct);
                }

                pendingRows.Add(valueRow);
            }

            foreach (var pendingRow in pendingRows) {
                sheetData.Append(pendingRow);
            }

            ClearHeaderCacheForPreparedAppend();
            int lastRow = startRow + rowCount + (includeHeaders ? 1 : 0) - 1;
            int lastColumn = startColumn + columnCount - 1;
            int dimensionMinRow = minExistingRow == int.MaxValue ? startRow : Math.Min(minExistingRow, startRow);
            int dimensionMinColumn = minExistingColumn == int.MaxValue ? startColumn : Math.Min(minExistingColumn, startColumn);
            int dimensionMaxRow = Math.Max(maxExistingRow, lastRow);
            int dimensionMaxColumn = Math.Max(maxExistingColumn, lastColumn);
            SetSheetDimensionReference(dimensionMinRow, dimensionMinColumn, dimensionMaxRow, dimensionMaxColumn);
            _requiresSavePreparation = false;
            return true;
        }

    }
}
