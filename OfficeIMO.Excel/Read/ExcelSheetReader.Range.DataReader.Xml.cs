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
        private sealed partial class ExcelXmlRangeDataReader : DbDataReader, IDataReaderFastValueSource {
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
            private readonly bool _trackCellPresence;
            private readonly CancellationToken _ct;
            private CancellationToken _activeReadCancellationToken;
            private readonly CultureInfo _culture;
            private readonly string[] _columnNames;
            private readonly Type[] _columnTypes;
            private readonly object?[] _currentValues;
            private readonly bool[] _currentValueLoaded;
            private readonly XmlDataReaderPrimitiveKind[] _currentPrimitiveKinds;
            private readonly double[] _currentDoubleValues;
            private readonly decimal[] _currentDecimalValues;
            private readonly DateTime[] _currentDateTimeValues;
            private readonly bool[] _currentBooleanValues;
            private readonly object?[] _blankRow;
            private readonly bool _hasRows;
            private Dictionary<int, object?[]>? _bufferedRows;
            private Dictionary<int, bool[]>? _bufferedCellPresence;
            private bool[]? _currentCellPresence;
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
            private bool _closed;
            private bool _disposed;

            IDataReaderFastValueSource? IDataReaderFastValueSource.FastValueSource =>
                _utf8Source != null ? this : null;

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
                int utf8SourceFirstColumn = 0,
                bool trackCellPresence = false,
                bool rowsAlreadyQualified = false) {
                _owner = owner;
                _firstRow = firstRow;
                _lastRow = lastRow;
                _firstColumn = firstColumn;
                _lastColumn = lastColumn;
                _fieldCount = fieldCount;
                _maximumBufferedCells = options.MaxDataReaderBufferedCells;
                _trackCellPresence = trackCellPresence;
                _ct = ct;
                _activeReadCancellationToken = ct;
                _culture = options.Culture;
                _nextLogicalRow = firstRow;
                _currentValues = new object?[fieldCount];
                _currentValueLoaded = new bool[fieldCount];
                _currentPrimitiveKinds = new XmlDataReaderPrimitiveKind[fieldCount];
                _currentDoubleValues = new double[fieldCount];
                _currentDecimalValues = new decimal[fieldCount];
                _currentDateTimeValues = new DateTime[fieldCount];
                _currentBooleanValues = new bool[fieldCount];
                _blankRow = new object?[fieldCount];

                if (preindexedUtf8Source != null) {
                    _utf8Source = preindexedUtf8Source;
                    _utf8SourceOrdinalOffset = firstColumn - utf8SourceFirstColumn;
                } else if (ExcelUtf8RangeRowSource.TryCreate(owner, firstRow, lastRow, firstColumn, fieldCount, ct, out var utf8Source)) {
                    _utf8Source = utf8Source;
                } else if (!owner._hasSdkWorksheetPart && !rowsAlreadyQualified) {
                    throw new XlsxTabularFastPathNotSupportedException(
                        $"Worksheet '{owner._sheetName}' requires the Open XML SDK fallback path.");
                } else {
                    _stream = owner.OpenDataReaderWorksheetStream(ct);
                    RewindWorksheetStream(_stream);
                    _reader = OpenWorksheetXmlReader(_stream);
                }

                try {
                    if (_trackCellPresence && _utf8Source == null) {
                        _currentCellPresence = new bool[fieldCount];
                    }
                    // XML rows may recur after a dense prefix. Establish ordering before
                    // publishing any values; unsorted input uses the existing cell budget.
                    if (_utf8Source == null && !rowsAlreadyQualified
                        && !owner.RowsAreSortedWithinRangeXmlFast(firstRow, lastRow, ct)) {
                        BufferRemainingRows();
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
                } catch {
                    _reader?.Dispose();
                    _stream.Dispose();
                    _utf8Source?.Dispose();
                    throw;
                }
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

            bool IDataReaderFastValueSource.TryGetUtf8Value(int ordinal, out ArraySegment<byte> value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }
                if (!_currentRowIsBlank
                    && _currentRow != null
                    && !_currentValueLoaded[ordinal]
                    && _utf8Source != null) {
                    return _utf8Source.TryGetUtf8Value(ordinal + _utf8SourceOrdinalOffset, out value);
                }

                value = default;
                return false;
            }

            bool IDataReaderFastValueSource.TryGetInt64(int ordinal, out long value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }
                if (_currentRowIsBlank || _currentRow == null) {
                    value = default;
                    return false;
                }
                if (TryGetUnloadedInt32(ordinal, out int integer)) {
                    value = integer;
                    return true;
                }
                if (!_currentValueLoaded[ordinal] && _utf8Source != null && !_owner._opt.NumericAsDecimal) {
                    _utf8Source.ReadValue(
                        ordinal + _utf8SourceOrdinalOffset,
                        XmlDataReaderTargetKind.Numeric,
                        out XmlDataReaderPrimitiveKind primitiveKind,
                        out double doubleValue,
                        out _,
                        out _,
                        out _,
                        out _,
                        out object? objectValue);
                    if (primitiveKind == XmlDataReaderPrimitiveKind.Double) {
                        value = Convert.ToInt64(doubleValue);
                        return true;
                    }
                    if (objectValue == null || objectValue == DBNull.Value) {
                        value = default;
                        return false;
                    }

                    value = Convert.ToInt64(objectValue, _culture);
                    return true;
                }
                if (IsDBNull(ordinal)) {
                    value = default;
                    return false;
                }

                value = GetInt64(ordinal);
                return true;
            }

            bool IDataReaderFastValueSource.TryGetDouble(int ordinal, out double value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }
                if (_currentRowIsBlank || _currentRow == null) {
                    value = default;
                    return false;
                }
                if (TryGetUnloadedNumber(ordinal, out value)) return true;
                if (!_currentValueLoaded[ordinal] && _utf8Source != null && !_owner._opt.NumericAsDecimal) {
                    _utf8Source.ReadValue(
                        ordinal + _utf8SourceOrdinalOffset,
                        XmlDataReaderTargetKind.Numeric,
                        out XmlDataReaderPrimitiveKind primitiveKind,
                        out double doubleValue,
                        out _,
                        out _,
                        out _,
                        out _,
                        out object? objectValue);
                    if (primitiveKind == XmlDataReaderPrimitiveKind.Double) {
                        value = doubleValue;
                        return true;
                    }
                    if (objectValue == null || objectValue == DBNull.Value) {
                        value = default;
                        return false;
                    }

                    value = Convert.ToDouble(objectValue, _culture);
                    return true;
                }
                if (IsDBNull(ordinal)) {
                    value = default;
                    return false;
                }

                value = GetDouble(ordinal);
                return true;
            }

            bool IDataReaderFastValueSource.TryGetDateTime(int ordinal, out DateTime value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }
                if (_currentRowIsBlank || _currentRow == null) {
                    value = default;
                    return false;
                }
                if (TryGetUnloadedDateTime(ordinal, out value)) return true;
                if (!_currentValueLoaded[ordinal] && _utf8Source != null) {
                    _utf8Source.ReadValue(
                        ordinal + _utf8SourceOrdinalOffset,
                        XmlDataReaderTargetKind.DateTime,
                        out XmlDataReaderPrimitiveKind primitiveKind,
                        out _,
                        out DateTime dateTimeValue,
                        out _,
                        out _,
                        out _,
                        out object? objectValue);
                    if (primitiveKind == XmlDataReaderPrimitiveKind.DateTime) {
                        value = dateTimeValue;
                        return true;
                    }
                    if (objectValue == null || objectValue == DBNull.Value) {
                        value = default;
                        return false;
                    }

                    value = objectValue is DateTime dateTime
                        ? dateTime
                        : Convert.ToDateTime(objectValue, _culture);
                    return true;
                }
                if (IsDBNull(ordinal)) {
                    value = default;
                    return false;
                }

                value = GetDateTime(ordinal);
                return true;
            }

            /// <inheritdoc />
            public override bool NextResult() => false;

            /// <inheritdoc />
            public override bool Read() => ReadCore(_ct);

            /// <inheritdoc />
            public override Task<bool> ReadAsync(CancellationToken cancellationToken) {
                try {
                    return Task.FromResult(ReadCore(cancellationToken));
                } catch (OperationCanceledException exception) when (exception.CancellationToken.IsCancellationRequested) {
                    return Task.FromCanceled<bool>(exception.CancellationToken);
                } catch (Exception exception) {
                    return Task.FromException<bool>(exception);
                }
            }

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
                _currentCellPresence = null;
                _bufferedCellPresence?.Clear();
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

            internal bool IsCellPresent(int ordinal) {
                if (_currentRowIsBlank) return false;
                if (_utf8Source != null) return _utf8Source.IsCellPresent(ordinal + _utf8SourceOrdinalOffset);
                EnsureCurrentValue(ordinal);
                return _currentCellPresence == null || _currentCellPresence[ordinal];
            }

            private bool TryReadLogicalRow(out object?[] row) {
                row = Array.Empty<object?>();
                if (_closed || _nextLogicalRow > _lastRow) {
                    return false;
                }

                ThrowIfReadCancellationRequested();
                if (_utf8Source != null) {
                    bool hasPhysicalRow = _utf8Source.SelectRow(_nextLogicalRow, _ct, _activeReadCancellationToken);
                    Array.Clear(_currentValueLoaded, 0, _currentValueLoaded.Length);
                    row = hasPhysicalRow ? _currentValues : _blankRow;
                    _currentRow = row;
                    _currentRowIsBlank = !hasPhysicalRow;
                    _currentRowActive = hasPhysicalRow;
                    _currentRowFinished = !hasPhysicalRow;
                    _nextLogicalRow++;
                    return true;
                }

                if (_bufferedRows != null) {
                    return TryReadBufferedLogicalRow(out row);
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
                    row = _blankRow;
                    _currentRow = row;
                    _currentRowIsBlank = true;
                    _nextLogicalRow++;
                    return true;
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

                    int rowIndex = ParsePositiveIntAttribute(ReadXmlReferenceAttribute(_reader).Text);
                    if (rowIndex <= 0) {
                        rowIndex = _owner.ResolveImplicitXmlRowIndex(_reader, _nextWorksheetRowIndex, _activeReadCancellationToken);
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
                    if (_bufferedCellPresence != null) {
                        _currentCellPresence = _bufferedCellPresence[_nextLogicalRow];
                        _bufferedCellPresence.Remove(_nextLogicalRow);
                    }
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
                if (_currentCellPresence != null) Array.Clear(_currentCellPresence, 0, _currentCellPresence.Length);
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

                if (_currentValueLoaded[ordinal]
                    && (_utf8Source != null || !_currentRowActive || _currentRowFinished)) {
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
                    if (!deferObjectMaterialization && _owner._opt.NumericAsDecimal
                        && _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.Double
                        && TryConvertExcelNumberToDecimal(_currentDoubleValues[ordinal], out _currentDecimalValues[ordinal])) {
                        _currentPrimitiveKinds[ordinal] = XmlDataReaderPrimitiveKind.Decimal;
                    }
                    return;
                }

                if (!_currentRowActive || _currentRowFinished) {
                    MarkCurrentValueMissing(ordinal);
                    return;
                }

                int targetColumn = _firstColumn + ordinal;
                // A later cell can repeat this column or appear before an earlier
                // one. Finish the bounded row cache before publishing a scalar so
                // every getter observes the same last-wins value as GetValues.
                ThrowIfReadCancellationRequested();
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

                    string? cellType = ReadXmlCellTypeAttribute(_reader);
                    if (_currentCellPresence != null) _currentCellPresence[columnOffset] = true;
                    if (columnIndex == targetColumn
                        && targetKind != XmlDataReaderTargetKind.None
                        && _owner.TryReadXmlCellPrimitiveForDataReader(
                            _reader,
                            cellType,
                            targetKind,
                            out XmlDataReaderPrimitiveKind primitiveKind,
                            out double doubleValue,
                            out decimal decimalValue,
                            out bool booleanValue,
                            out object? objectValue)) {
                        _currentValues[columnOffset] = objectValue;
                        _currentPrimitiveKinds[columnOffset] = primitiveKind;
                        _currentDoubleValues[columnOffset] = doubleValue;
                        _currentDecimalValues[columnOffset] = decimalValue;
                        _currentBooleanValues[columnOffset] = booleanValue;
                    } else {
                        _currentValues[columnOffset] = _owner.ReadXmlCellValue(_reader, cellType, preserveDateSerial: true);
                        _currentPrimitiveKinds[columnOffset] = XmlDataReaderPrimitiveKind.None;
                    }

                    _currentValueLoaded[columnOffset] = true;
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
                    ThrowIfReadCancellationRequested();
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

                        if (_currentCellPresence != null) _currentCellPresence[columnOffset] = true;
                        _currentValues[columnOffset] = _owner.ReadXmlCellValue(_reader, ReadXmlCellTypeAttribute(_reader), preserveDateSerial: true);
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
                if (_trackCellPresence) _bufferedCellPresence ??= new Dictionary<int, bool[]>();
                if (_hasPendingRow) {
                    ReadBufferedRowValues(_pendingRowIndex);
                    _hasPendingRow = false;
                }

                while (_reader.Read()) {
                    ThrowIfReadCancellationRequested();

                    if (_reader.NodeType != XmlNodeType.Element || _reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(ReadXmlReferenceAttribute(_reader).Text);
                    if (rowIndex <= 0) {
                        rowIndex = _owner.ResolveImplicitXmlRowIndex(_reader, _nextWorksheetRowIndex, _activeReadCancellationToken);
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

                    ReadBufferedRowValues(rowIndex);
                }
            }

            private void ReadBufferedRowValues(int rowIndex) {
                if (rowIndex < _nextLogicalRow || rowIndex > _lastRow) {
                    SkipXmlElement(_reader, "row");
                    return;
                }

                if (!_bufferedRows!.TryGetValue(rowIndex, out object?[]? values)) {
                    if ((long)_bufferedRows.Count + 1L > _maximumBufferedCells / _fieldCount) {
                        throw new InvalidDataException($"Range data-reader buffering exceeds {nameof(ExcelReadOptions.MaxDataReaderBufferedCells)}.");
                    }
                    values = new object?[_fieldCount];
                    _bufferedRows.Add(rowIndex, values);
                    _bufferedCellPresence?.Add(rowIndex, new bool[_fieldCount]);
                }

                // Apply only the present cells to the logical row. A new empty
                // array for every fragment would discard earlier omitted cells.
                _owner.ReadXmlRowIntoChunk(
                    _reader,
                    new[] { values },
                    rowIndex,
                    rowIndex,
                    _firstColumn,
                    _lastColumn,
                    _activeReadCancellationToken,
                    preserveDateSerial: true,
                    cellPresence: _bufferedCellPresence?[rowIndex]);
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
                headerValues != null && ordinal < headerValues.Length
                    ? (headerValues[ordinal] is ExcelDataReaderDateSerial serial ? serial.Materialize() : headerValues[ordinal])?.ToString() : null;

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
