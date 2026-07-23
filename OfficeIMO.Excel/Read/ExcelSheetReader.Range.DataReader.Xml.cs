#nullable enable

using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Data-reader projections for <see cref="ExcelSheetReader"/> ranges.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private sealed class ExcelXmlRangeDataReader : DbDataReader {
            private const string IsReadOnlyColumn = "IsReadOnly";
            private const string IsRowVersionColumn = "IsRowVersion";
            private const string IsAutoIncrementColumn = "IsAutoIncrement";
            private const string BaseCatalogNameColumn = "BaseCatalogName";

            private readonly ExcelSheetReader _owner;
            private readonly Stream _stream = Stream.Null;
            private readonly XmlReader _reader = null!;
            private readonly ExcelUtf8RangeRowSource? _utf8Source;
            private readonly int _firstRow;
            private readonly int _lastRow;
            private readonly int _firstColumn;
            private readonly int _lastColumn;
            private readonly int _fieldCount;
            private readonly long _maximumBufferedCells;
            private readonly CancellationToken _ct;
            private readonly string[] _columnNames;
            private readonly Type[] _columnTypes;
            private readonly object?[] _currentValues;
            private readonly bool[] _currentValueLoaded;
            private readonly XmlDataReaderPrimitiveKind[] _currentPrimitiveKinds;
            private readonly double[] _currentDoubleValues;
            private readonly DateTime[] _currentDateTimeValues;
            private readonly bool[] _currentBooleanValues;
            private readonly object?[] _blankRow;
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
                CancellationToken ct) {
                _owner = owner;
                _firstRow = firstRow;
                _lastRow = lastRow;
                _firstColumn = firstColumn;
                _lastColumn = lastColumn;
                _fieldCount = fieldCount;
                _maximumBufferedCells = options.MaxDataReaderBufferedCells;
                _ct = ct;
                _nextLogicalRow = firstRow;
                _currentValues = new object?[fieldCount];
                _currentValueLoaded = new bool[fieldCount];
                _currentPrimitiveKinds = new XmlDataReaderPrimitiveKind[fieldCount];
                _currentDoubleValues = new double[fieldCount];
                _currentDateTimeValues = new DateTime[fieldCount];
                _currentBooleanValues = new bool[fieldCount];
                _blankRow = new object?[fieldCount];

                if (ExcelUtf8RangeRowSource.TryCreate(owner, firstRow, lastRow, firstColumn, fieldCount, ct, out var utf8Source)) {
                    _utf8Source = utf8Source;
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
            public override bool HasRows => !_closed && _nextLogicalRow <= _lastRow;

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
                return value is bool boolean ? boolean : Convert.ToBoolean(value, CultureInfo.InvariantCulture);
            }

            /// <inheritdoc />
            public override byte GetByte(int ordinal) => Convert.ToByte(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

            /// <inheritdoc />
            public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) =>
                throw new NotSupportedException("Excel range fields are exposed as scalar values.");

            /// <inheritdoc />
            public override char GetChar(int ordinal) => Convert.ToChar(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

            /// <inheritdoc />
            public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) {
                string value = Convert.ToString(GetValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty;
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
                return value is DateTime dateTime ? dateTime : Convert.ToDateTime(value, CultureInfo.InvariantCulture);
            }

            /// <inheritdoc />
            public override decimal GetDecimal(int ordinal) => Convert.ToDecimal(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

            /// <inheritdoc />
            public override double GetDouble(int ordinal) {
                EnsureOpenRow();
                EnsureCurrentValue(ordinal, XmlDataReaderTargetKind.Double);
                if (IsCurrentStreamingRow && _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.Double) {
                    return _currentDoubleValues[ordinal];
                }

                object value = GetNonDbNullValue(ordinal);
                return value is double number ? number : Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }

            /// <inheritdoc />
            [UnconditionalSuppressMessage("Trimming", "IL2063", Justification = "Excel reader column types are closed scalar conversion tokens; OfficeIMO never activates or reflects over their public members.")]
            [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
            public override Type GetFieldType(int ordinal) => _columnTypes[ordinal];

            /// <inheritdoc />
            public override float GetFloat(int ordinal) => Convert.ToSingle(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

            /// <inheritdoc />
            public override Guid GetGuid(int ordinal) => (Guid)GetNonDbNullValue(ordinal);

            /// <inheritdoc />
            public override short GetInt16(int ordinal) => Convert.ToInt16(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

            /// <inheritdoc />
            public override int GetInt32(int ordinal) {
                EnsureOpenRow();
                EnsureCurrentValue(ordinal, XmlDataReaderTargetKind.Int32);
                if (IsCurrentStreamingRow && _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.Double) {
                    return ConvertDataReaderInt32(_currentDoubleValues[ordinal]);
                }

                object value = GetNonDbNullValue(ordinal);
                return ConvertDataReaderInt32(value);
            }

            /// <inheritdoc />
            public override long GetInt64(int ordinal) => Convert.ToInt64(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

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
                return value is string text ? text : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
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
            public override bool Read() {
                if (_closed) {
                    return false;
                }

                if (TryReadLogicalRow(out var row)) {
                    _currentRow = row;
                    return true;
                }

                _currentRow = null;
                return false;
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
            public override DataTable GetSchemaTable() {
                var schema = new DataTable("SchemaTable");
                schema.Columns.Add(SchemaTableColumn.ColumnName, typeof(string));
                schema.Columns.Add(SchemaTableColumn.ColumnOrdinal, typeof(int));
                schema.Columns.Add(SchemaTableColumn.ColumnSize, typeof(int));
                schema.Columns.Add(SchemaTableColumn.NumericPrecision, typeof(short));
                schema.Columns.Add(SchemaTableColumn.NumericScale, typeof(short));
                schema.Columns.Add(SchemaTableColumn.DataType, typeof(Type));
                schema.Columns.Add(SchemaTableColumn.ProviderType, typeof(int));
                schema.Columns.Add(SchemaTableColumn.IsLong, typeof(bool));
                schema.Columns.Add(SchemaTableColumn.AllowDBNull, typeof(bool));
                schema.Columns.Add(IsReadOnlyColumn, typeof(bool));
                schema.Columns.Add(IsRowVersionColumn, typeof(bool));
                schema.Columns.Add(SchemaTableColumn.IsUnique, typeof(bool));
                schema.Columns.Add(SchemaTableColumn.IsKey, typeof(bool));
                schema.Columns.Add(IsAutoIncrementColumn, typeof(bool));
                schema.Columns.Add(SchemaTableColumn.BaseSchemaName, typeof(string));
                schema.Columns.Add(BaseCatalogNameColumn, typeof(string));
                schema.Columns.Add(SchemaTableColumn.BaseTableName, typeof(string));
                schema.Columns.Add(SchemaTableColumn.BaseColumnName, typeof(string));

                for (int i = 0; i < _fieldCount; i++) {
                    var row = schema.NewRow();
                    row[SchemaTableColumn.ColumnName] = _columnNames[i];
                    row[SchemaTableColumn.ColumnOrdinal] = i;
                    row[SchemaTableColumn.ColumnSize] = -1;
                    row[SchemaTableColumn.NumericPrecision] = DBNull.Value;
                    row[SchemaTableColumn.NumericScale] = DBNull.Value;
                    row[SchemaTableColumn.DataType] = _columnTypes[i];
                    row[SchemaTableColumn.ProviderType] = 0;
                    row[SchemaTableColumn.IsLong] = false;
                    row[SchemaTableColumn.AllowDBNull] = true;
                    row[IsReadOnlyColumn] = true;
                    row[IsRowVersionColumn] = false;
                    row[SchemaTableColumn.IsUnique] = false;
                    row[SchemaTableColumn.IsKey] = false;
                    row[IsAutoIncrementColumn] = false;
                    row[SchemaTableColumn.BaseSchemaName] = DBNull.Value;
                    row[BaseCatalogNameColumn] = DBNull.Value;
                    row[SchemaTableColumn.BaseTableName] = DBNull.Value;
                    row[SchemaTableColumn.BaseColumnName] = _columnNames[i];
                    schema.Rows.Add(row);
                }

                return schema;
            }

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

                _ct.ThrowIfCancellationRequested();
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
                    _rowsAreSorted ??= _owner.RowsAreSortedWithinRangeXmlFast(_firstRow, _lastRow, _ct);
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
                    if (_ct.CanBeCanceled) {
                        _ct.ThrowIfCancellationRequested();
                    }

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
                        ordinal,
                        targetKind,
                        out _currentPrimitiveKinds[ordinal],
                        out _currentDoubleValues[ordinal],
                        out _currentDateTimeValues[ordinal],
                        out _currentBooleanValues[ordinal],
                        out _,
                        out _currentValues[ordinal]);
                    _currentValueLoaded[ordinal] = true;
                    return;
                }

                if (!_currentRowActive || _currentRowFinished) {
                    MarkCurrentValueMissing(ordinal);
                    return;
                }

                int targetColumn = _firstColumn + ordinal;
                while (_reader.Read()) {
                    if (_ct.CanBeCanceled) {
                        _ct.ThrowIfCancellationRequested();
                    }

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
                        if (_ct.CanBeCanceled) {
                            _ct.ThrowIfCancellationRequested();
                        }

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
                    if (_ct.CanBeCanceled) {
                        _ct.ThrowIfCancellationRequested();
                    }

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
                    _owner.ReadXmlRowIntoChunk(_reader, rowSlot, rowIndex, rowIndex, _firstColumn, _lastColumn, _ct);
                    StoreBufferedRow(rowIndex, values);
                }
            }

            private object?[] ReadPendingRowValues() {
                var values = new object?[_fieldCount];
                var rowSlot = new[] { values };
                _owner.ReadXmlRowIntoChunk(_reader, rowSlot, _pendingRowIndex, _pendingRowIndex, _firstColumn, _lastColumn, _ct);
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

            private object? MaterializeCurrentValue(int ordinal) {
                if (!IsCurrentStreamingRow || _currentPrimitiveKinds[ordinal] == XmlDataReaderPrimitiveKind.None) {
                    return _currentRow![ordinal];
                }

                object value = _currentPrimitiveKinds[ordinal] switch {
                    XmlDataReaderPrimitiveKind.Double => _currentDoubleValues[ordinal],
                    XmlDataReaderPrimitiveKind.DateTime => _currentDateTimeValues[ordinal],
                    XmlDataReaderPrimitiveKind.Boolean => BoxBoolean(_currentBooleanValues[ordinal]),
                    _ => _currentRow![ordinal]!
                };
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
