#nullable enable

using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Data-reader projections for <see cref="ExcelSheetReader"/> ranges.
    /// </summary>
    internal sealed partial class ExcelSheetReader {
        private sealed class ExcelXmlRangeDataReader : DbDataReader {
            private readonly ExcelSheetReader _owner;
            private readonly Stream _stream = Stream.Null;
            private readonly XmlReader _reader = null!;
            private readonly ExcelUtf8RangeRowSource? _utf8Source;
            private readonly int _utf8SourceOrdinalOffset;
            private readonly int _firstRow;
            private readonly int _lastRow;
            private readonly int _firstColumn;
            private readonly int _lastColumn;
            private readonly int _fieldCount;
            private readonly long _maximumBufferedCells;
            private readonly CancellationToken _ct;
            private CancellationToken _activeReadCancellationToken;
            private readonly CultureInfo _culture;
            private readonly string[] _columnNames;
            private readonly Type[] _columnTypes;
            private readonly object?[] _currentValues;
            private readonly bool[] _currentValueLoaded;
            private readonly XmlDataReaderPrimitiveKind[] _currentPrimitiveKinds;
            private readonly double[] _currentDoubleValues;
            private readonly DateTime[] _currentDateTimeValues;
            private readonly bool[] _currentBooleanValues;
            private readonly object?[] _blankRow;
            private readonly bool _hasRows;
            private Dictionary<int, object?[]>? _bufferedRows;
            private Dictionary<string, int>? _ordinals;
            private object?[]? _currentRow;
            private int _nextLogicalRow;
            private int _nextWorksheetRowIndex = 1;
            private int _pendingRowIndex;
            private int _currentRowDepth;
            private int _currentNextCellColumnIndex = 1;
            private bool _hasPendingRow;
            private bool _currentRowActive;
            private bool _currentRowFinished;
            private bool _currentRowIsBlank;
            private bool? _rowsAreSorted;
            private bool _closed;
            private bool _disposed;

            internal ExcelXmlRangeDataReader(
                ExcelSheetReader owner,
                int firstRow,
                int firstColumn,
                int lastRow,
                int lastColumn,
                int fieldCount,
                bool headersInFirstRow,
                ExcelReadOptions options,
                CancellationToken ct,
                ExcelUtf8RangeRowSource? preindexedUtf8Source = null,
                int utf8SourceFirstColumn = 0) {
                _owner = owner;
                _firstRow = firstRow;
                _lastRow = lastRow;
                _firstColumn = firstColumn;
                _lastColumn = lastColumn;
                _fieldCount = fieldCount;
                _maximumBufferedCells = options.MaxDataReaderBufferedCells;
                _ct = ct;
                _activeReadCancellationToken = ct;
                _culture = options.Culture;
                _nextLogicalRow = firstRow;
                _currentValues = new object?[fieldCount];
                _currentValueLoaded = new bool[fieldCount];
                _currentPrimitiveKinds = new XmlDataReaderPrimitiveKind[fieldCount];
                _currentDoubleValues = new double[fieldCount];
                _currentDateTimeValues = new DateTime[fieldCount];
                _currentBooleanValues = new bool[fieldCount];
                _blankRow = new object?[fieldCount];

                if (preindexedUtf8Source != null) {
                    _utf8Source = preindexedUtf8Source;
                    _utf8SourceOrdinalOffset = firstColumn - utf8SourceFirstColumn;
                } else if (ExcelUtf8RangeRowSource.TryCreate(owner, firstRow, lastRow, firstColumn, fieldCount, ct, out var utf8Source)) {
                    _utf8Source = utf8Source;
                } else if (!owner._hasSdkWorksheetPart) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        $"Worksheet '{owner._sheetName}' requires the Open XML SDK fallback path.");
                } else {
                    _stream = owner._wsPart.GetStream(FileMode.Open, FileAccess.Read);
                    RewindWorksheetStream(_stream);
                    _reader = OpenWorksheetXmlReader(_stream);
                }

                object?[]? headerValues = null;
                if (headersInFirstRow) {
                    if (TryReadLogicalRow(out headerValues)) {
                        MaterializeAllCurrentRowValues();
                        headerValues = _currentRow;
                    }
                }

                _columnNames = headersInFirstRow
                    ? ExcelHeaderNameHelper.BuildUniqueHeaders(fieldCount, c => GetHeaderText(headerValues, c), options.NormalizeHeaders)
                    : CreateGeneratedColumnNames(fieldCount);
                _columnTypes = CreateObjectColumnTypes(fieldCount);
                _hasRows = _nextLogicalRow <= _lastRow;
                _currentRow = null;
            }

            /// <inheritdoc />
            public override object this[int ordinal] => GetValue(ordinal);

            /// <inheritdoc />
            public override object this[string name] => GetValue(GetOrdinal(name));

            /// <inheritdoc />
            public override int Depth => 0;

            /// <inheritdoc />
            public override int FieldCount => _fieldCount;

            /// <inheritdoc />
            public override bool HasRows => !_closed && _hasRows;

            /// <inheritdoc />
            public override bool IsClosed => _closed;

            /// <inheritdoc />
            public override int RecordsAffected => -1;

            /// <inheritdoc />
            public override bool GetBoolean(int ordinal) {
                EnsureOpenRow();
                EnsureCurrentValue(ordinal, XmlDataReaderTargetKind.Boolean);
                if (IsCurrentStreamingRow && _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.Boolean) {
                    return _currentBooleanValues[ordinal];
                }

                object value = GetNonDbNullValue(ordinal);
                return value is bool boolean ? boolean : Convert.ToBoolean(value, _culture);
            }

            /// <inheritdoc />
            public override byte GetByte(int ordinal) => TryGetPrimitiveDouble(ordinal, out double value)
                ? Convert.ToByte(value)
                : Convert.ToByte(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) =>
                throw new NotSupportedException("Excel range fields are exposed as scalar values.");

            /// <inheritdoc />
            public override char GetChar(int ordinal) => Convert.ToChar(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) {
                string value = Convert.ToString(GetValue(ordinal), _culture) ?? string.Empty;
                if (buffer == null) {
                    return value.Length;
                }

                if (dataOffset >= value.Length || length == 0) {
                    return 0;
                }

                int offset = (int)dataOffset;
                int count = Math.Min(length, value.Length - offset);
                if (count <= 0) {
                    return 0;
                }

                value.CopyTo(offset, buffer, bufferOffset, count);
                return count;
            }

            /// <inheritdoc />
            public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;

            /// <inheritdoc />
            public override DateTime GetDateTime(int ordinal) {
                EnsureOpenRow();
                EnsureCurrentValue(ordinal, XmlDataReaderTargetKind.DateTime);
                if (IsCurrentStreamingRow && _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.DateTime) {
                    return _currentDateTimeValues[ordinal];
                }

                object value = GetNonDbNullValue(ordinal);
                return value is DateTime dateTime ? dateTime : Convert.ToDateTime(value, _culture);
            }

            /// <inheritdoc />
            public override decimal GetDecimal(int ordinal) => TryGetPrimitiveDouble(ordinal, out double value)
                ? Convert.ToDecimal(value)
                : Convert.ToDecimal(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override double GetDouble(int ordinal) {
                return TryGetPrimitiveDouble(ordinal, out double value)
                    ? value
                    : Convert.ToDouble(GetNonDbNullValue(ordinal), _culture);
            }

            /// <inheritdoc />
            [UnconditionalSuppressMessage("Trimming", "IL2063", Justification = "Excel reader column types are closed scalar conversion tokens; OfficeIMO never activates or reflects over their public members.")]
            [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
            public override Type GetFieldType(int ordinal) => _columnTypes[ordinal];

            /// <inheritdoc />
            public override float GetFloat(int ordinal) => TryGetPrimitiveDouble(ordinal, out double value)
                ? (float)value
                : Convert.ToSingle(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override Guid GetGuid(int ordinal) {
                object value = GetNonDbNullValue(ordinal);
                return value is Guid guid ? guid : Guid.Parse(Convert.ToString(value, _culture)!);
            }

            /// <inheritdoc />
            public override short GetInt16(int ordinal) => TryGetPrimitiveDouble(ordinal, out double value)
                ? Convert.ToInt16(value)
                : Convert.ToInt16(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override int GetInt32(int ordinal) {
                return TryGetPrimitiveDouble(ordinal, out double value)
                    ? ConvertDataReaderInt32(value)
                    : ConvertDataReaderInt32(GetNonDbNullValue(ordinal), _culture);
            }

            /// <inheritdoc />
            public override long GetInt64(int ordinal) => TryGetPrimitiveDouble(ordinal, out double value)
                ? Convert.ToInt64(value)
                : Convert.ToInt64(GetNonDbNullValue(ordinal), _culture);

            /// <inheritdoc />
            public override string GetName(int ordinal) => _columnNames[ordinal];

            /// <inheritdoc />
            public override int GetOrdinal(string name) {
                _ordinals ??= CreateOrdinalMap(_columnNames);
                if (_ordinals.TryGetValue(name, out int ordinal)) {
                    return ordinal;
                }

                throw new IndexOutOfRangeException(name);
            }

            /// <inheritdoc />
            public override string GetString(int ordinal) {
                object value = GetNonDbNullValue(ordinal);
                return value is string text ? text : Convert.ToString(value, _culture) ?? string.Empty;
            }

            /// <inheritdoc />
            public override object GetValue(int ordinal) {
                EnsureOpenRow();
                EnsureCurrentValue(ordinal);
                object? value = MaterializeCurrentValue(ordinal);
                return ToDataReaderValue(value);
            }

            /// <inheritdoc />
            public override int GetValues(object[] values) {
                EnsureOpenRow();
                MaterializeAllCurrentRowValues();
                MaterializeAllPrimitiveCurrentValues();
                return CopyDataReaderValues(_currentRow!, _fieldCount, values);
            }

            /// <inheritdoc />
            public override bool IsDBNull(int ordinal) => GetValue(ordinal) == DBNull.Value;

            /// <inheritdoc />
            public override bool NextResult() => false;

            /// <inheritdoc />
            public override bool Read() => ReadCore(_ct);

            /// <inheritdoc />
            public override Task<bool> ReadAsync(CancellationToken cancellationToken) =>
                Task.Run(() => ReadCore(cancellationToken), cancellationToken);

            private bool ReadCore(CancellationToken cancellationToken) {
                _ct.ThrowIfCancellationRequested();
                cancellationToken.ThrowIfCancellationRequested();
                _activeReadCancellationToken = cancellationToken;
                try {
                    if (_closed) {
                        return false;
                    }

                    if (TryReadLogicalRow(out var row)) {
                        _currentRow = row;
                        return true;
                    }

                    _currentRow = null;
                    return false;
                } finally {
                    _activeReadCancellationToken = _ct;
                }
            }

            /// <inheritdoc />
            public override void Close() {
                if (_closed) {
                    return;
                }

                _closed = true;
                _currentRow = null;
                _utf8Source?.Dispose();
                _reader?.Dispose();
                if (!ReferenceEquals(_stream, Stream.Null)) {
                    _stream.Dispose();
                }
            }

            /// <inheritdoc />
            [UnconditionalSuppressMessage("Trimming", "IL2111", Justification = "The schema table stores Type values as data and does not reflect over Type.TypeInitializer or other Type members.")]
            public override DataTable GetSchemaTable() =>
                ExcelDataReaderSchemaTable.Create(_fieldCount, GetName, GetFieldType);

            /// <inheritdoc />
            public override IEnumerator GetEnumerator() {
                while (Read()) {
                    yield return this;
                }
            }

            /// <inheritdoc />
            protected override void Dispose(bool disposing) {
                if (disposing && !_disposed) {
                    _disposed = true;
                    Close();
                }

                base.Dispose(disposing);
            }

            private bool TryReadLogicalRow(out object?[] row) {
                row = Array.Empty<object?>();
                if (_closed || _nextLogicalRow > _lastRow) {
                    return false;
                }

                ThrowIfReadCancellationRequested();
                if (_utf8Source != null) {
                    Array.Clear(_currentValueLoaded, 0, _currentValueLoaded.Length);
                    bool hasPhysicalRow = _utf8Source.SelectRow(_nextLogicalRow);
                    row = hasPhysicalRow ? _currentValues : _blankRow;
                    _currentRow = row;
                    _currentRowIsBlank = !hasPhysicalRow;
                    _currentRowActive = hasPhysicalRow;
                    _currentRowFinished = !hasPhysicalRow;
                    _nextLogicalRow++;
                    return true;
                }

                FinishCurrentRow();
                EnsurePendingRow();
                if (_hasPendingRow && _pendingRowIndex == _nextLogicalRow) {
                    BeginPendingRow();
                    row = _currentValues;
                    _hasPendingRow = false;
                    _nextLogicalRow++;
                    return true;
                }

                if (_hasPendingRow && _pendingRowIndex > _nextLogicalRow) {
                    _rowsAreSorted ??= _owner.RowsAreSortedWithinRangeXmlFast(
                        _firstRow,
                        _lastRow,
                        _activeReadCancellationToken);
                    if (_rowsAreSorted.Value) {
                        row = _blankRow;
                        _currentRow = row;
                        _currentRowIsBlank = true;
                        _nextLogicalRow++;
                        return true;
                    }

                    BufferRemainingRows();
                    return TryReadBufferedLogicalRow(out row);
                }

                if (_bufferedRows != null) {
                    return TryReadBufferedLogicalRow(out row);
                }

                row = _blankRow;
                _currentRow = row;
                _currentRowIsBlank = true;
                _nextLogicalRow++;
                return true;
            }

            private void EnsurePendingRow() {
                if (_hasPendingRow || _nextLogicalRow > _lastRow) {
                    return;
                }

                while (_reader.Read()) {
                    ThrowIfReadCancellationRequested();

                    if (_reader.NodeType != XmlNodeType.Element || _reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(_reader.GetAttribute("r"));
                    if (rowIndex <= 0) {
                        rowIndex = _nextWorksheetRowIndex;
                    }

                    _nextWorksheetRowIndex = rowIndex + 1;
                    if (rowIndex < _firstRow) {
                        SkipXmlElement(_reader, "row");
                        continue;
                    }

                    if (rowIndex < _nextLogicalRow) {
                        SkipXmlElement(_reader, "row");
                        continue;
                    }

                    _pendingRowIndex = rowIndex;
                    _hasPendingRow = true;
                    return;
                }
            }

            private bool TryReadBufferedLogicalRow(out object?[] row) {
                row = Array.Empty<object?>();
                if (_closed || _nextLogicalRow > _lastRow) {
                    return false;
                }

                if (_bufferedRows != null && _bufferedRows.TryGetValue(_nextLogicalRow, out var bufferedRow)) {
                    row = bufferedRow;
                    _bufferedRows.Remove(_nextLogicalRow);
                    _currentRowIsBlank = false;
                } else {
                    row = _blankRow;
                    _currentRowIsBlank = true;
                }

                _currentRow = row;
                _currentRowActive = false;
                _currentRowFinished = true;
                _nextLogicalRow++;
                return true;
            }

            private void BeginPendingRow() {
                Array.Clear(_currentValueLoaded, 0, _currentValueLoaded.Length);
                _currentRow = _currentValues;
                _currentRowDepth = _reader.Depth;
                _currentNextCellColumnIndex = 1;
                _currentRowIsBlank = false;
                _currentRowActive = !_reader.IsEmptyElement;
                _currentRowFinished = _reader.IsEmptyElement;
            }

            private void FinishCurrentRow() {
                if (_utf8Source != null) {
                    _currentRowActive = false;
                    _currentRowFinished = true;
                    _currentRow = null;
                    _currentRowIsBlank = false;
                    return;
                }

                if (_currentRowActive && !_currentRowFinished) {
                    SkipXmlElementContent(_reader, _currentRowDepth);
                }

                _currentRowActive = false;
                _currentRowFinished = true;
                _currentRow = null;
                _currentRowIsBlank = false;
            }

            private void EnsureCurrentValue(int ordinal, XmlDataReaderTargetKind targetKind = XmlDataReaderTargetKind.None) {
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }

                if (_currentRowIsBlank || _currentRow == null) {
                    return;
                }

                if (_currentValueLoaded[ordinal]) {
                    return;
                }

                if (_utf8Source != null) {
                    _utf8Source.ReadValue(
                        ordinal + _utf8SourceOrdinalOffset,
                        targetKind,
                        out _currentPrimitiveKinds[ordinal],
                        out _currentDoubleValues[ordinal],
                        out _currentDateTimeValues[ordinal],
                        out _currentBooleanValues[ordinal],
                        out _,
                        out bool deferObjectMaterialization,
                        out _currentValues[ordinal]);
                    _currentValueLoaded[ordinal] = !deferObjectMaterialization;
                    return;
                }

                if (!_currentRowActive || _currentRowFinished) {
                    MarkCurrentValueMissing(ordinal);
                    return;
                }

                int targetColumn = _firstColumn + ordinal;
                while (_reader.Read()) {
                    ThrowIfReadCancellationRequested();

                    if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == _currentRowDepth && _reader.LocalName == "row") {
                        _currentRowActive = false;
                        _currentRowFinished = true;
                        break;
                    }

                    if (_reader.NodeType != XmlNodeType.Element || _reader.LocalName != "c") {
                        continue;
                    }

                    int columnIndex = GetXmlCellColumnIndex(_reader, ref _currentNextCellColumnIndex);
                    if (columnIndex <= 0) {
                        SkipXmlElement(_reader, "c");
                        continue;
                    }

                    if (columnIndex < _firstColumn || columnIndex > _lastColumn) {
                        SkipXmlElement(_reader, "c");
                        continue;
                    }

                    int columnOffset = columnIndex - _firstColumn;
                    if ((uint)columnOffset >= (uint)_fieldCount) {
                        SkipXmlElement(_reader, "c");
                        continue;
                    }

                    string? cellType = _reader.GetAttribute("t");
                    if (columnIndex == targetColumn
                        && targetKind != XmlDataReaderTargetKind.None
                        && _owner.TryReadXmlCellPrimitiveForDataReader(
                            _reader,
                            cellType,
                            targetKind,
                            out XmlDataReaderPrimitiveKind primitiveKind,
                            out double doubleValue,
                            out DateTime dateTimeValue,
                            out bool booleanValue,
                            out object? objectValue)) {
                        _currentValues[columnOffset] = objectValue;
                        _currentPrimitiveKinds[columnOffset] = primitiveKind;
                        _currentDoubleValues[columnOffset] = doubleValue;
                        _currentDateTimeValues[columnOffset] = dateTimeValue;
                        _currentBooleanValues[columnOffset] = booleanValue;
                    } else {
                        _currentValues[columnOffset] = _owner.ReadXmlCellValue(_reader, cellType);
                        _currentPrimitiveKinds[columnOffset] = XmlDataReaderPrimitiveKind.None;
                    }

                    _currentValueLoaded[columnOffset] = true;

                    if (columnIndex == targetColumn) {
                        return;
                    }
                }

                if (!_currentValueLoaded[ordinal]) {
                    MarkCurrentValueMissing(ordinal);
                }
            }

            private void MaterializeAllCurrentRowValues() {
                if (_currentRowIsBlank || _currentRow == null) {
                    return;
                }

                if (_utf8Source != null) {
                    for (int i = 0; i < _fieldCount; i++) {
                        EnsureCurrentValue(i);
                    }
                    return;
                }

                if (_currentRowActive && !_currentRowFinished) {
                    while (_reader.Read()) {
                        ThrowIfReadCancellationRequested();

                        if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == _currentRowDepth && _reader.LocalName == "row") {
                            _currentRowActive = false;
                            _currentRowFinished = true;
                            break;
                        }

                        if (_reader.NodeType != XmlNodeType.Element || _reader.LocalName != "c") {
                            continue;
                        }

                        int columnIndex = GetXmlCellColumnIndex(_reader, ref _currentNextCellColumnIndex);
                        if (columnIndex <= 0) {
                            SkipXmlElement(_reader, "c");
                            continue;
                        }

                        if (columnIndex < _firstColumn || columnIndex > _lastColumn) {
                            SkipXmlElement(_reader, "c");
                            continue;
                        }

                        int columnOffset = columnIndex - _firstColumn;
                        if ((uint)columnOffset >= (uint)_fieldCount) {
                            SkipXmlElement(_reader, "c");
                            continue;
                        }

                        _currentValues[columnOffset] = _owner.ReadXmlCellValue(_reader, _reader.GetAttribute("t"));
                        _currentPrimitiveKinds[columnOffset] = XmlDataReaderPrimitiveKind.None;
                        _currentValueLoaded[columnOffset] = true;
                    }
                }

                for (int i = 0; i < _currentValueLoaded.Length; i++) {
                    if (!_currentValueLoaded[i]) {
                        MarkCurrentValueMissing(i);
                    }
                }
            }

            private void MarkCurrentValueMissing(int ordinal) {
                _currentValues[ordinal] = null;
                _currentPrimitiveKinds[ordinal] = XmlDataReaderPrimitiveKind.None;
                _currentValueLoaded[ordinal] = true;
            }

            private void BufferRemainingRows() {
                _bufferedRows ??= new Dictionary<int, object?[]>();
                if (_hasPendingRow) {
                    StoreBufferedRow(_pendingRowIndex, ReadPendingRowValues());
                    _hasPendingRow = false;
                }

                while (_reader.Read()) {
                    ThrowIfReadCancellationRequested();

                    if (_reader.NodeType != XmlNodeType.Element || _reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(_reader.GetAttribute("r"));
                    if (rowIndex <= 0) {
                        rowIndex = _nextWorksheetRowIndex;
                    }

                    _nextWorksheetRowIndex = rowIndex + 1;
                    if (rowIndex < _nextLogicalRow) {
                        SkipXmlElement(_reader, "row");
                        continue;
                    }

                    if (rowIndex > _lastRow) {
                        SkipXmlElement(_reader, "row");
                        continue;
                    }

                    var values = new object?[_fieldCount];
                    var rowSlot = new[] { values };
                    _owner.ReadXmlRowIntoChunk(
                        _reader,
                        rowSlot,
                        rowIndex,
                        rowIndex,
                        _firstColumn,
                        _lastColumn,
                        _activeReadCancellationToken);
                    StoreBufferedRow(rowIndex, values);
                }
            }

            private object?[] ReadPendingRowValues() {
                var values = new object?[_fieldCount];
                var rowSlot = new[] { values };
                _owner.ReadXmlRowIntoChunk(
                    _reader,
                    rowSlot,
                    _pendingRowIndex,
                    _pendingRowIndex,
                    _firstColumn,
                    _lastColumn,
                    _activeReadCancellationToken);
                return values;
            }

            private void StoreBufferedRow(int rowIndex, object?[] values) {
                if (rowIndex < _nextLogicalRow || rowIndex > _lastRow) {
                    return;
                }

                if (!_bufferedRows!.ContainsKey(rowIndex) &&
                    (long)_bufferedRows.Count + 1L > _maximumBufferedCells / _fieldCount) {
                    throw new InvalidDataException($"Range data-reader buffering exceeds {nameof(ExcelReadOptions.MaxDataReaderBufferedCells)}.");
                }

                var copy = new object?[_fieldCount];
                Array.Copy(values, copy, Math.Min(values.Length, copy.Length));
                _bufferedRows![rowIndex] = copy;
            }

            private object GetNonDbNullValue(int ordinal) {
                EnsureOpenRow();
                EnsureCurrentValue(ordinal);
                object? value = MaterializeCurrentValue(ordinal);
                if (value == null || value == DBNull.Value) {
                    throw new InvalidCastException($"Column '{GetName(ordinal)}' contains DBNull.");
                }

                return value;
            }

            private bool TryGetPrimitiveDouble(int ordinal, out double value) {
                EnsureOpenRow();
                EnsureCurrentValue(ordinal, XmlDataReaderTargetKind.Numeric);
                if (IsCurrentStreamingRow
                    && _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.Double) {
                    value = _currentDoubleValues[ordinal];
                    return true;
                }

                value = 0;
                return false;
            }

            private object? MaterializeCurrentValue(int ordinal) {
                if (!IsCurrentStreamingRow || _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.None) {
                    return _currentRow![ordinal];
                }

                object value = _currentValues[ordinal] ?? (_currentPrimitiveKinds[ordinal] switch {
                    XmlDataReaderPrimitiveKind.Double => _currentDoubleValues[ordinal],
                    XmlDataReaderPrimitiveKind.DateTime => _currentDateTimeValues[ordinal],
                    XmlDataReaderPrimitiveKind.Boolean => BoxBoolean(_currentBooleanValues[ordinal]),
                    _ => _currentRow![ordinal]!
                });
                _currentValues[ordinal] = value;
                _currentPrimitiveKinds[ordinal] = XmlDataReaderPrimitiveKind.None;
                return value;
            }

            private void MaterializeAllPrimitiveCurrentValues() {
                if (!IsCurrentStreamingRow) {
                    return;
                }

                for (int i = 0; i < _currentPrimitiveKinds.Length; i++) {
                    if (_currentPrimitiveKinds[i] != XmlDataReaderPrimitiveKind.None) {
                        _ = MaterializeCurrentValue(i);
                    }
                }
            }

            private bool IsCurrentStreamingRow => ReferenceEquals(_currentRow, _currentValues);

            private void EnsureOpenRow() {
                if (_closed) {
                    throw new InvalidOperationException("The reader is closed.");
                }

                if (_currentRow == null) {
                    throw new InvalidOperationException("The reader is not positioned on a row.");
                }
            }

            private void ThrowIfReadCancellationRequested() {
                _ct.ThrowIfCancellationRequested();
                if (_activeReadCancellationToken != _ct) {
                    _activeReadCancellationToken.ThrowIfCancellationRequested();
                }
            }

            private static string? GetHeaderText(object?[]? headerValues, int ordinal) =>
                headerValues != null && ordinal < headerValues.Length ? headerValues[ordinal]?.ToString() : null;

            private static string[] CreateGeneratedColumnNames(int fieldCount) {
                var names = new string[fieldCount];
                for (int i = 0; i < names.Length; i++) {
                    names[i] = $"Column{i + 1}";
                }

                return names;
            }

            private static Type[] CreateObjectColumnTypes(int fieldCount) {
                var types = new Type[fieldCount];
                for (int i = 0; i < types.Length; i++) {
                    types[i] = typeof(object);
                }

                return types;
            }

        }
    }
}
