using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int DirectSequentialCellWriteLimit = 16;

        /// <summary>
        /// Writes multiple cell values efficiently, using parallelization when beneficial.
        /// </summary>
        /// <param name="cells">Collection of cell coordinates and values.</param>
        /// <param name="mode">Optional execution mode override.</param>
        /// <param name="ct">Cancellation token.</param>
        /// <remarks>
        /// This is the canonical API for batch cell writes. Use this in place of the older
        /// <see cref="SetCellValues(IEnumerable{ValueTuple{int, int, object}}, ExecutionMode?, CancellationToken)"/>
        /// method, which will be removed in a future release.
        /// </remarks>
        public void CellValues(IEnumerable<(int Row, int Column, object Value)> cells, ExecutionMode? mode = null, CancellationToken ct = default) {
            if (cells is null) {
                throw new ArgumentNullException(nameof(cells));
            }
            var list = cells as IList<(int Row, int Column, object Value)> ?? cells.ToList();
            if (list.Count == 0) return;

            // Single cell: trivially sequential
            if (list.Count == 1) {
                var single = list[0];
                CellValue(single.Row, single.Column, single.Value);
                return;
            }

            if (list.Count > DirectSequentialCellWriteLimit && TryApplyPlainCellsByAppendingRows(list, ct)) {
                return;
            }

            // Prepared buffers for parallel scenario
            var prepared = new (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[list.Count];
            var ssPlanner = new SharedStringPlanner();

            ExecuteWithPolicy(
                opName: "CellValues",
                itemCount: list.Count,
                overrideMode: mode,
                sequentialCore: () => {
                    if (list.Count <= DirectSequentialCellWriteLimit) {
                        for (int i = 0; i < list.Count; i++) {
                            ct.ThrowIfCancellationRequested();
                            var (r, c, v) = list[i];
                            CellValueCore(r, c, v);
                        }

                        return;
                    }

                    // Sequential path - keep the fast prepared/apply writer so row-major
                    // batches can append rows instead of falling back to GetCell per cell.
                    for (int i = 0; i < list.Count; i++) {
                        ct.ThrowIfCancellationRequested();
                        var (r, c, v) = list[i];
                        var (val, type) = CoerceForCellNoDom(v, ssPlanner);
                        prepared[i] = (r, c, val!, type!);
                    }

                    ssPlanner.ApplyAndFixup(prepared, _excelDocument);
                    ApplyPreparedCells(prepared, list);
                },
                computeParallel: () => {
                    // Parallel compute phase - prepare values without DOM mutation
                    Parallel.For(0, list.Count, new ParallelOptions {
                        CancellationToken = ct,
                        MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
                    }, i => {
                        var (r, c, obj) = list[i];
                        var (val, type) = CoerceForCellNoDom(obj, ssPlanner);
                        prepared[i] = (r, c, val!, type!);
                    });
                },
                applySequential: () => {
                    // Apply phase - first fix shared strings, then write all values to DOM
                    ssPlanner.ApplyAndFixup(prepared, _excelDocument);
                    ApplyPreparedCells(prepared, list);
                },
                ct: ct
            );
        }

        /// <summary>
        /// Compute-only coercion for parallel scenarios. Does not mutate DOM.
        /// Uses <see cref="SharedStringPlanner"/> for string values.
        /// </summary>
        private (CellValue cellValue, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> dataType) CoerceForCellNoDom(object? value, SharedStringPlanner planner) {
            var dateTimeOffsetStrategy = _excelDocument.DateTimeOffsetWriteStrategy;
            var (cellValue, cellType) = CoerceValueHelper.Coerce(
                value,
                s => {
                    var sanitized = planner.Note(s);
                    return new CellValue(sanitized);
                },
                dateTimeOffsetStrategy);
            return (cellValue, new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cellType));
        }

        private bool TryApplyPlainCellsByAppendingRows(IList<(int Row, int Column, object Value)> source, CancellationToken ct) {
            bool applied = false;
            System.Threading.ReaderWriterLockSlim? lck = _excelDocument._lock;
            if (lck == null) {
                try { lck = _excelDocument.EnsureLock(); } catch { lck = null; }
            }

            Locking.ExecuteWrite(lck, () => applied = TryApplyPlainCellsByAppendingRowsCore(source, ct));
            return applied;
        }

        private bool TryApplyPlainCellsByAppendingRowsCore(IList<(int Row, int Column, object Value)> source, CancellationToken ct) {
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
            var appendedRows = new List<OpenXmlElement>(Math.Max(1, source.Count / Math.Max(maxColumn, 1)));
            Row? row = null;
            List<OpenXmlElement>? rowCells = null;
            int rowIndex = 0;
            string rowReference = string.Empty;

            for (int i = 0; i < source.Count; i++) {
                ct.ThrowIfCancellationRequested();
                var item = source[i];

                if (item.Row != rowIndex) {
                    if (row != null) {
                        row.Append(rowCells!);
                        appendedRows.Add(row);
                    }

                    rowIndex = item.Row;
                    rowReference = rowIndex.ToString(CultureInfo.InvariantCulture);
                    row = new Row { RowIndex = (uint)rowIndex };
                    rowCells = new List<OpenXmlElement>(Math.Min(maxColumn, 16));
                }

                var (cellValue, dataType) = CoercePlainAppendValue(item.Value, ref sharedStringIndexes, useDirectStringCells);
                rowCells!.Add(new Cell {
                    CellReference = columnNames[item.Column] + rowReference,
                    CellValue = cellValue,
                    DataType = dataType
                });
            }

            if (row != null) {
                row.Append(rowCells!);
                appendedRows.Add(row);
            }

            sheetData.Append(appendedRows);
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
            lock (_headerMapLock) {
                _headerMapCache = null;
                _headerMapSourceA1 = null;
            }
        }

        private bool TryGetPlainAppendLayout(
            IList<(int Row, int Column, object Value)> source,
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
            var dimensions = worksheet.Elements<SheetDimension>().ToList();
            SheetDimension? dimension = dimensions.FirstOrDefault();
            foreach (var extraDimension in dimensions.Skip(1).ToList()) {
                extraDimension.Remove();
            }

            string start = A1.CellReference(minRow, minColumn);
            string end = A1.CellReference(maxRow, maxColumn);
            string reference = start == end ? start : start + ":" + end;
            if (dimension == null) {
                worksheet.InsertAt(new SheetDimension { Reference = reference }, 0);
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

        private (CellValue cellValue, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> dataType) CoercePlainAppendValue(
            object? value,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool useDirectStringCells) {
            (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) = value switch {
                null => CoerceValueHelper.HandleEmptyString(),
                DBNull => CoerceValueHelper.HandleEmptyString(),
                string text => useDirectStringCells
                    ? (CreatePlainAppendStringValue(text), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                    : (CreatePlainAppendSharedStringValue(text, ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString),
                double number => CoerceValueHelper.HandleNumber(number),
                float number => CoerceValueHelper.HandleNumber(Convert.ToDouble(number)),
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

            return (cellValue, new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cellType));
        }

        private static CellValue CreatePlainAppendStringValue(string text) {
            CoerceValueHelper.ValidateSharedStringLength(text, nameof(text));
            return new CellValue(Utilities.ExcelSanitizer.SanitizeString(text));
        }

        private CellValue CreatePlainAppendSharedStringValue(string text, ref Dictionary<string, int>? sharedStringIndexes) {
            string sanitized = Utilities.ExcelSanitizer.SanitizeString(text);
            sharedStringIndexes ??= new Dictionary<string, int>(StringComparer.Ordinal);
            if (!sharedStringIndexes.TryGetValue(sanitized, out int index)) {
                index = _excelDocument.GetSharedStringIndex(sanitized);
                sharedStringIndexes[sanitized] = index;
            }

            return new CellValue(index.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Obsolete. Use <see cref="CellValues(IEnumerable{ValueTuple{int, int, object}}, ExecutionMode?, CancellationToken)"/> instead.
        /// </summary>
        [Obsolete("Use CellValues(...) instead.")]
        public void SetCellValues(IEnumerable<(int Row, int Column, object Value)> cells, ExecutionMode? mode = null, CancellationToken ct = default) {
            CellValues(cells, mode, ct);
        }

        private void ApplyPreparedCells(
            (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[] prepared,
            IList<(int Row, int Column, object Value)> source) {
            if (TryApplyPreparedCellsByAppendingRows(prepared, source)) {
                return;
            }

            var writer = new BatchCellWriter(this);

            for (int i = 0; i < prepared.Length; i++) {
                var p = prepared[i];
                var originalValue = source[i].Value;
                var cell = writer.GetOrCreateCell(p.Row, p.Col);
                cell.CellValue = p.Val;
                cell.DataType = p.Type;
                ApplyAutomaticCellFormatting(cell, originalValue, p.Type);
            }

            ClearHeaderCache();
        }

        private bool TryApplyPreparedCellsByAppendingRows(
            (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[] prepared,
            IList<(int Row, int Column, object Value)> source) {
            if (prepared.Length != source.Count) {
                return false;
            }

            if (prepared.Length == 0) {
                ClearHeaderCache();
                return true;
            }

            int firstRow = prepared[0].Row;
            int currentRow = 0;
            int currentColumn = 0;
            int maxColumn = 0;

            for (int i = 0; i < prepared.Length; i++) {
                var p = prepared[i];
                if (p.Row <= 0 || p.Col <= 0 || p.Col > A1.MaxColumns || p.Row < currentRow) {
                    return false;
                }

                if (p.Row != currentRow) {
                    currentRow = p.Row;
                    currentColumn = 0;
                }

                if (p.Col <= currentColumn) {
                    return false;
                }

                currentColumn = p.Col;
                if (p.Col > maxColumn) {
                    maxColumn = p.Col;
                }
            }

            var sheetData = GetOrCreateSheetData();
            foreach (var existingRow in sheetData.Elements<Row>()) {
                if (existingRow.RowIndex == null) {
                    return false;
                }

                if (existingRow.RowIndex != null && existingRow.RowIndex.Value >= (uint)firstRow) {
                    return false;
                }
            }

            bool needsAutomaticFormatting = false;
            for (int i = 0; i < source.Count; i++) {
                if (RequiresAutomaticCellFormatting(source[i].Value, prepared[i].Type)) {
                    needsAutomaticFormatting = true;
                    break;
                }
            }

            var columnNames = new string[maxColumn + 1];
            for (int column = 1; column <= maxColumn; column++) {
                columnNames[column] = GetColumnName(column);
            }

            var baseStyleIndexes = needsAutomaticFormatting
                ? GetAppendBaseStyleIndexes(sheetData, firstRow, maxColumn)
                : null;
            Row? row = null;
            int rowIndex = 0;
            string rowReference = string.Empty;
            Dictionary<uint, uint>? appendedDateStyleIndexes = null;
            Dictionary<uint, uint>? appendedDurationStyleIndexes = null;
            Dictionary<uint, uint>? appendedWrapStyleIndexes = null;
            for (int i = 0; i < prepared.Length; i++) {
                var p = prepared[i];
                if (p.Row != rowIndex) {
                    rowIndex = p.Row;
                    rowReference = rowIndex.ToString(CultureInfo.InvariantCulture);
                    row = new Row { RowIndex = (uint)rowIndex };
                    sheetData.Append(row);
                }

                var cell = new Cell {
                    CellReference = columnNames[p.Col] + rowReference,
                    CellValue = p.Val,
                    DataType = p.Type
                };

                row!.Append(cell);
                if (needsAutomaticFormatting) {
                    ApplyAutomaticCellFormattingForAppendedCell(
                        cell,
                        source[i].Value,
                        p.Type,
                        baseStyleIndexes![p.Col] ?? 0U,
                        ref appendedDateStyleIndexes,
                        ref appendedDurationStyleIndexes,
                        ref appendedWrapStyleIndexes);
                }
            }

            ClearHeaderCache();
            return true;
        }

        private static uint?[] GetAppendBaseStyleIndexes(SheetData sheetData, int firstRow, int maxColumn) {
            var baseStyleIndexes = new uint?[maxColumn + 1];
            var baseStyleRows = new int[maxColumn + 1];

            foreach (var existingRow in sheetData.Elements<Row>()) {
                if (existingRow.RowIndex == null) {
                    continue;
                }

                int existingRowIndex = (int)existingRow.RowIndex.Value;
                if (existingRowIndex >= firstRow) {
                    continue;
                }

                foreach (var existingCell in existingRow.Elements<Cell>()) {
                    if (existingCell.CellReference == null || existingCell.StyleIndex == null) {
                        continue;
                    }

                    int columnIndex = A1.ParseColumnIndexFromCellReference(existingCell.CellReference.Value);
                    if (columnIndex <= 0 || columnIndex > maxColumn) {
                        continue;
                    }

                    if (existingRowIndex >= baseStyleRows[columnIndex]) {
                        baseStyleRows[columnIndex] = existingRowIndex;
                        baseStyleIndexes[columnIndex] = existingCell.StyleIndex.Value;
                    }
                }
            }

            return baseStyleIndexes;
        }

        private sealed class BatchCellWriter {
            private readonly ExcelSheet _sheet;
            private readonly SheetData _sheetData;
            private readonly Dictionary<int, BatchRowState> _rows;
            private Row? _lastRow;
            private int _lastRowIndex;

            internal BatchCellWriter(ExcelSheet sheet) {
                _sheet = sheet;
                _sheetData = sheet.GetOrCreateSheetData();
                _rows = new Dictionary<int, BatchRowState>();

                foreach (var row in _sheetData.Elements<Row>()) {
                    if (row.RowIndex == null) {
                        continue;
                    }

                    _rows[(int)row.RowIndex.Value] = new BatchRowState(row);
                    if (row.RowIndex.Value >= _lastRowIndex) {
                        _lastRowIndex = (int)row.RowIndex.Value;
                        _lastRow = row;
                    }
                }
            }

            internal Cell GetOrCreateCell(int rowIndex, int columnIndex) {
                if (!_rows.TryGetValue(rowIndex, out BatchRowState? rowState)) {
                    var row = GetOrCreateRow(rowIndex);
                    rowState = new BatchRowState(row);
                    _rows[rowIndex] = rowState;
                }

                return rowState.GetOrCreateCell(columnIndex, rowIndex);
            }

            private Row GetOrCreateRow(int rowIndex) {
                if (_lastRow != null && rowIndex > _lastRowIndex) {
                    var appended = new Row { RowIndex = (uint)rowIndex };
                    _sheetData.Append(appended);
                    _lastRow = appended;
                    _lastRowIndex = rowIndex;
                    return appended;
                }

                var row = _sheet.GetOrCreateRowElement(_sheetData, rowIndex);
                if (row.RowIndex != null && row.RowIndex.Value >= _lastRowIndex) {
                    _lastRow = row;
                    _lastRowIndex = (int)row.RowIndex.Value;
                }

                return row;
            }

            private sealed class BatchRowState {
                private readonly Row _row;
                private readonly Dictionary<int, Cell> _cells;
                private Cell? _lastCell;
                private int _lastColumnIndex;

                internal BatchRowState(Row row) {
                    _row = row;
                    _cells = new Dictionary<int, Cell>();

                    foreach (var cell in row.Elements<Cell>()) {
                        var reference = cell.CellReference?.Value;
                        if (string.IsNullOrEmpty(reference)) {
                            continue;
                        }

                        int columnIndex = GetColumnIndex(reference!);
                        _cells[columnIndex] = cell;

                        if (columnIndex >= _lastColumnIndex) {
                            _lastColumnIndex = columnIndex;
                            _lastCell = cell;
                        }
                    }
                }

                internal Cell GetOrCreateCell(int columnIndex, int rowIndex) {
                    if (_cells.TryGetValue(columnIndex, out Cell? existing)) {
                        return existing;
                    }

                    string cellReference = A1.CellReference(rowIndex, columnIndex);
                    var cell = new Cell { CellReference = cellReference };

                    if (_lastCell == null) {
                        var firstCell = _row.Elements<Cell>().FirstOrDefault();
                        if (firstCell != null) {
                            _row.InsertBefore(cell, firstCell);
                        } else {
                            _row.Append(cell);
                        }
                    } else if (columnIndex > _lastColumnIndex) {
                        _row.InsertAfter(cell, _lastCell);
                    } else {
                        Cell? insertAfter = null;
                        foreach (var existingCell in _row.Elements<Cell>()) {
                            var existingReference = existingCell.CellReference?.Value;
                            if (string.IsNullOrEmpty(existingReference)) {
                                continue;
                            }

                            int existingColumnIndex = GetColumnIndex(existingReference!);
                            if (existingColumnIndex > columnIndex) {
                                _row.InsertBefore(cell, existingCell);
                                _cells[columnIndex] = cell;
                                return cell;
                            }

                            insertAfter = existingCell;
                        }

                        if (insertAfter != null) {
                            _row.InsertAfter(cell, insertAfter);
                        } else {
                            _row.Append(cell);
                        }
                    }

                    _cells[columnIndex] = cell;
                    if (columnIndex >= _lastColumnIndex) {
                        _lastColumnIndex = columnIndex;
                        _lastCell = cell;
                    }

                    return cell;
                }
            }
        }
    }
}
