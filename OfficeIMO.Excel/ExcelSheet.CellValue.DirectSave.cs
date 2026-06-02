using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        private sealed class CellValueDirectSaveBuffer {
            private readonly List<object?[]> _rows = new();
            private readonly List<int> _filledCounts = new();
            private Type[] _columnTypes = Array.Empty<Type>();
            private int _columnCount;
            private int _lockedColumnCount;
            private int _lastRow;
            private int _lastColumn;

            internal IReadOnlyList<object?[]> Rows => _rows;

            internal Type[] ColumnTypes => _columnTypes;

            internal int ColumnCount => _columnCount;

            internal int RowCount => _rows.Count;

            internal int CellCount => _filledCounts.Count == 0 ? 0 : ((_rows.Count - 1) * _columnCount) + _filledCounts[_filledCounts.Count - 1];

            internal bool TryAdd(int row, int column, object? value) {
                if (row <= 0 || column <= 0) {
                    return false;
                }

                if (_rows.Count == 0) {
                    if (row != 1 || column != 1) {
                        return false;
                    }

                    EnsureColumnCount(1);
                    _rows.Add(new object?[1]);
                    _filledCounts.Add(0);
                    SetValue(0, 1, value);
                    _lastRow = 1;
                    _lastColumn = 1;
                    return true;
                }

                if (row == _lastRow && column == _lastColumn + 1) {
                    if (_lockedColumnCount > 0 && column > _lockedColumnCount) {
                        return false;
                    }

                    EnsureColumnCount(column);
                    SetValue(row - 1, column, value);
                    _lastColumn = column;
                    return true;
                }

                if (row == _lastRow + 1 && column == 1) {
                    if (_lockedColumnCount == 0) {
                        _lockedColumnCount = _columnCount;
                    }

                    if (_lastColumn != _lockedColumnCount) {
                        return false;
                    }

                    _rows.Add(new object?[_columnCount]);
                    _filledCounts.Add(0);
                    SetValue(row - 1, 1, value);
                    _lastRow = row;
                    _lastColumn = 1;
                    return true;
                }

                return false;
            }

            internal bool IsComplete {
                get {
                    if (_rows.Count == 0 || _columnCount == 0) {
                        return false;
                    }

                    for (int i = 0; i < _filledCounts.Count; i++) {
                        if (_filledCounts[i] != _columnCount) {
                            return false;
                        }
                    }

                    return true;
                }
            }

            internal IEnumerable<(int Row, int Column, object? Value)> EnumerateWrittenCells() {
                for (int row = 0; row < _rows.Count; row++) {
                    object?[] values = _rows[row];
                    int count = _filledCounts[row];
                    for (int column = 0; column < count; column++) {
                        yield return (row + 1, column + 1, values[column]);
                    }
                }
            }

            private void EnsureColumnCount(int column) {
                if (column <= _columnCount) {
                    return;
                }

                int oldColumnCount = _columnCount;
                _columnCount = column;
                Array.Resize(ref _columnTypes, _columnCount);
                for (int i = oldColumnCount; i < _columnTypes.Length; i++) {
                    _columnTypes[i] = typeof(object);
                }

                for (int i = 0; i < _rows.Count; i++) {
                    object?[] row = _rows[i];
                    Array.Resize(ref row, _columnCount);
                    _rows[i] = row;
                }
            }

            private void SetValue(int rowIndex, int column, object? value) {
                _rows[rowIndex][column - 1] = value;
                _filledCounts[rowIndex] = column;
                if (value == null || value == DBNull.Value) {
                    return;
                }

                Type valueType = value.GetType();
                int columnIndex = column - 1;
                Type currentType = _columnTypes[columnIndex];
                _columnTypes[columnIndex] = currentType == typeof(object) || currentType == valueType
                    ? valueType
                    : typeof(object);
            }
        }

        // Core implementation: single source of truth (no locks here)
        private void CellValueCore(int row, int column, object? value) {
            MaterializeDeferredDataSetImportIfNeeded();
            CellValueCoreNoMaterialize(row, column, value);
        }

        private bool TrySetPendingDirectCellValue(int row, int column, object? value) {
            if (_isBatchOperation || Locking.IsNoLock) {
                return TrySetPendingDirectCellValueLocked(row, column, value);
            }

            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                return TrySetPendingDirectCellValueLocked(row, column, value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        private bool TrySetPendingDirectCellValueLocked(int row, int column, object? value) {
            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (!EnablePendingDirectCellValueBuffer
                || _materializingPendingCellValueDirectSaveBuffer
                || _disablePendingCellValueDirectSaveBuffer
                || _excelDocument.HasDeferredDirectDataSetImport
                || _excelDocument.HasDirectDataSetFastSaveState
                || !TryPreparePendingDirectCellValue(value, out object? directValue)
                || (_pendingCellValueDirectSaveBuffer == null && _hasCellValueDomWrites)
                || (_pendingCellValueDirectSaveBuffer == null && !CanRegisterDirectTabularSaveCandidate(1, 1, Math.Max(1, column)))) {
                return false;
            }

            return TrySetPendingDirectCellValueCore(row, column, directValue);
        }

        private bool TrySetPendingDirectCellFormula(int row, int column, string formula) {
            if (_isBatchOperation || Locking.IsNoLock) {
                return TrySetPendingDirectCellFormulaLocked(row, column, formula);
            }

            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                return TrySetPendingDirectCellFormulaLocked(row, column, formula);
            } finally {
                lck.ExitWriteLock();
            }
        }

        private bool TrySetPendingDirectCellFormulaLocked(int row, int column, string formula) {
            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (!EnablePendingDirectCellValueBuffer
                || _materializingPendingCellValueDirectSaveBuffer
                || _excelDocument.HasDeferredDirectDataSetImport
                || _excelDocument.HasDirectDataSetFastSaveState
                || (_pendingCellValueDirectSaveBuffer == null && _hasCellValueDomWrites)
                || (_pendingCellValueDirectSaveBuffer == null && !CanRegisterDirectTabularSaveCandidate(1, 1, Math.Max(1, column)))) {
                return false;
            }

            var directValue = new DirectFormulaCellValue(Utilities.ExcelSanitizer.SanitizeFormula(formula));
            return TrySetPendingDirectCellValueCore(row, column, directValue);
        }

        private bool TrySetPendingDirectCellValueCore(int row, int column, object? value) {
            if (!_excelDocument.TryReservePendingDirectCellValueSheet(this)) {
                return false;
            }

            int currentThreadId = Environment.CurrentManagedThreadId;
            if (_pendingCellValueDirectSaveBuffer == null) {
                _pendingCellValueDirectSaveThreadId = currentThreadId;
            } else if (_pendingCellValueDirectSaveThreadId != currentThreadId) {
                _disablePendingCellValueDirectSaveBuffer = true;
                MaterializePendingDirectCellValues();
                return false;
            }

            var buffer = _pendingCellValueDirectSaveBuffer ??= new CellValueDirectSaveBuffer();
            if (!buffer.TryAdd(row, column, value)) {
                MaterializePendingDirectCellValues();
                return false;
            }

            if (MirrorPendingDirectCellValueBufferToWorksheet || buffer.CellCount < PendingDirectCellValueMinimumCellCount) {
                ApplyPendingDirectCellValueToDom(row, column, value);
            }

            if (!_excelDocument.IsPackageDirty) {
                _excelDocument.MarkPackageDirty();
            }

            return true;
        }

        private void ApplyPendingDirectCellValueToDom(int row, int column, object? value) {
            if (value is DirectFormulaCellValue formula) {
                CellFormulaCore(row, column, formula.Formula);
                return;
            }

            CellValueCoreNoMaterialize(row, column, value);
        }

        private bool TryPreparePendingDirectCellValue(object? value, out object? directValue) {
            if (value == null || value == DBNull.Value) {
                directValue = null;
                return true;
            }

            switch (value) {
                case string text:
                    CoerceValueHelper.ValidateSharedStringLength(text, nameof(value));
                    directValue = text;
                    return text.IndexOf('\n') < 0 && text.IndexOf('\r') < 0;
                case double:
                case float:
                case decimal:
                case int:
                case long:
                case short:
                case uint:
                case ulong:
                case ushort:
                case byte:
                case sbyte:
                case bool:
                    directValue = value;
                    return true;
                case DateTime dateTime:
                    _ = dateTime.ToOADate();
                    directValue = dateTime;
                    return true;
                case DateTimeOffset dateTimeOffset:
                    if (TryPrepareDateTimeOffsetPendingDirectCellValue(dateTimeOffset, out DateTime convertedDateTime)) {
                        directValue = convertedDateTime;
                        return true;
                    }

                    directValue = value;
                    return false;
                case TimeSpan:
                    directValue = value;
                    return true;
#if NET6_0_OR_GREATER
                case DateOnly dateOnly:
                    _ = dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate();
                    directValue = dateOnly;
                    return true;
                case TimeOnly:
                    directValue = value;
                    return true;
#endif
                default:
                    directValue = value;
                    return false;
            }
        }

        private bool TryPrepareDateTimeOffsetPendingDirectCellValue(DateTimeOffset value, out DateTime converted) {
            try {
                converted = _excelDocument.DateTimeOffsetWriteStrategy(value);
            } catch (Exception ex) {
                throw new InvalidOperationException("The configured DateTimeOffset write strategy threw an exception.", ex);
            }

            if (value.UtcDateTime < CellValueExcelMinimumSupportedDate) {
                return false;
            }

            try {
                _ = converted.ToOADate();
                return true;
            } catch (ArgumentException) {
                return false;
            } catch (OverflowException) {
                return false;
            }
        }

        internal void MaterializePendingDirectCellValues() {
            var buffer = _pendingCellValueDirectSaveBuffer;
            if (buffer == null) {
                return;
            }

            _pendingCellValueDirectSaveBuffer = null;
            _excelDocument.ClearPendingDirectCellValueSheet(this);

            _materializingPendingCellValueDirectSaveBuffer = true;
            try {
                foreach (var cell in buffer.EnumerateWrittenCells()) {
                    ApplyPendingDirectCellValueToDom(cell.Row, cell.Column, cell.Value);
                }
            } finally {
                _materializingPendingCellValueDirectSaveBuffer = false;
            }
        }

        internal bool TryPromotePendingDirectCellValuesToSaveCandidate() {
            var buffer = _pendingCellValueDirectSaveBuffer;
            if (buffer == null) {
                return false;
            }

            if (!buffer.IsComplete
                || buffer.ColumnCount <= 0
                || buffer.RowCount <= 0
                || buffer.CellCount < PendingDirectCellValueMinimumCellCount) {
                return false;
            }

            var columnNames = new string[buffer.ColumnCount];
            for (int i = 0; i < columnNames.Length; i++) {
                columnNames[i] = "Column" + (i + 1).ToString(CultureInfo.InvariantCulture);
            }

            string range = A1.CellReference(1, 1) + ":" + A1.CellReference(buffer.RowCount, buffer.ColumnCount);
            bool registered = _excelDocument.RegisterDeferredDirectTabularSaveCandidate(
                this,
                "Cells",
                columnNames,
                buffer.ColumnTypes,
                buffer.Rows,
                includeHeaders: false,
                range,
                useCellValueNumberFormats: true,
                replacingPendingDirectCellValues: true);
            if (registered) {
                _pendingCellValueDirectSaveBuffer = null;
                _excelDocument.ClearPendingDirectCellValueSheet(this);
            }

            return registered;
        }
    }
}
