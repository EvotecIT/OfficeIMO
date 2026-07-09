#nullable enable

using System.Collections;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Data-reader projections for <see cref="ExcelSheetReader"/> ranges.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Creates a forward-only <see cref="IDataReader"/> over a rectangular range without
        /// materializing the full range. The reader buffers at most <paramref name="schemaSampleRows"/>
        /// data rows to infer column types, then continues from the worksheet stream.
        /// </summary>
        /// <param name="a1Range">A rectangular A1 range to read.</param>
        /// <param name="headersInFirstRow">When true, the first row supplies column names.</param>
        /// <param name="chunkRows">Number of worksheet rows requested from each streaming chunk.</param>
        /// <param name="schemaSampleRows">Maximum number of data rows buffered for schema inference.</param>
        /// <param name="mode">Optional execution mode used by the range stream.</param>
        /// <param name="ct">Cancellation token observed while streaming.</param>
        /// <returns>A forward-only reader that must be disposed after use.</returns>
        public IDataReader ReadRangeAsDataReader(
            string a1Range,
            bool headersInFirstRow = true,
            int chunkRows = 1024,
            int schemaSampleRows = 1024,
            OfficeIMO.Excel.ExecutionMode? mode = null,
            CancellationToken ct = default) {
            if (schemaSampleRows < 0) {
                throw new ArgumentOutOfRangeException(nameof(schemaSampleRows), "Schema sample row count cannot be negative.");
            }

            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) {
                throw new ArgumentException($"Invalid range '{a1Range}'.", nameof(a1Range));
            }

            int cols = c2 - c1 + 1;
            IEnumerable<RangeChunk> chunks = ReadRangeStream(a1Range, chunkRows, mode, ct);
            return new ExcelRangeDataReader(chunks, r1, r2, cols, headersInFirstRow, schemaSampleRows, _opt, ct);
        }

        private sealed class ExcelRangeDataReader : DbDataReader {
            private const string IsReadOnlyColumn = "IsReadOnly";
            private const string IsRowVersionColumn = "IsRowVersion";
            private const string IsAutoIncrementColumn = "IsAutoIncrement";
            private const string BaseCatalogNameColumn = "BaseCatalogName";

            private readonly IEnumerator<RangeChunk> _chunks;
            private readonly int _lastRow;
            private readonly int _fieldCount;
            private readonly CancellationToken _ct;
            private readonly string[] _columnNames;
            private readonly Type[] _columnTypes;
            private readonly Dictionary<string, int> _ordinals;
            private readonly List<object?[]> _prefetchedRows;
            private RangeChunk? _currentChunk;
            private object?[]? _currentRow;
            private int _nextRow;
            private int _prefetchedIndex;
            private bool _closed;
            private bool _disposed;

            internal ExcelRangeDataReader(
                IEnumerable<RangeChunk> chunks,
                int firstRow,
                int lastRow,
                int fieldCount,
                bool headersInFirstRow,
                int schemaSampleRows,
                ExcelReadOptions options,
                CancellationToken ct) {
                _chunks = chunks.GetEnumerator();
                _nextRow = firstRow;
                _lastRow = lastRow;
                _fieldCount = fieldCount;
                _ct = ct;

                object?[]? headerValues = null;
                if (headersInFirstRow) {
                    TryReadLogicalRow(out headerValues);
                }

                _columnNames = headersInFirstRow
                    ? ExcelHeaderNameHelper.BuildUniqueHeaders(fieldCount, c => GetHeaderText(headerValues, c), options.NormalizeHeaders)
                    : CreateGeneratedColumnNames(fieldCount);

                _prefetchedRows = new List<object?[]>(Math.Min(schemaSampleRows, 1024));
                while (_prefetchedRows.Count < schemaSampleRows && TryReadLogicalRow(out var row)) {
                    _prefetchedRows.Add(row);
                }

                _columnTypes = options.InferDataTableColumnTypes
                    ? InferColumnTypes(_prefetchedRows, fieldCount)
                    : CreateObjectColumnTypes(fieldCount);

                _ordinals = new Dictionary<string, int>(fieldCount, StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < _columnNames.Length; i++) {
                    _ordinals[_columnNames[i]] = i;
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
            public override bool HasRows => !_closed && (_prefetchedIndex < _prefetchedRows.Count || _nextRow <= _lastRow);

            /// <inheritdoc />
            public override bool IsClosed => _closed;

            /// <inheritdoc />
            public override int RecordsAffected => -1;

            /// <inheritdoc />
            public override bool GetBoolean(int ordinal) => (bool)GetNonDbNullValue(ordinal);

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
            public override DateTime GetDateTime(int ordinal) => (DateTime)GetNonDbNullValue(ordinal);

            /// <inheritdoc />
            public override decimal GetDecimal(int ordinal) => Convert.ToDecimal(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

            /// <inheritdoc />
            public override double GetDouble(int ordinal) => (double)GetNonDbNullValue(ordinal);

            /// <inheritdoc />
            public override Type GetFieldType(int ordinal) => _columnTypes[ordinal];

            /// <inheritdoc />
            public override float GetFloat(int ordinal) => Convert.ToSingle(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

            /// <inheritdoc />
            public override Guid GetGuid(int ordinal) => (Guid)GetNonDbNullValue(ordinal);

            /// <inheritdoc />
            public override short GetInt16(int ordinal) => Convert.ToInt16(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

            /// <inheritdoc />
            public override int GetInt32(int ordinal) => Convert.ToInt32(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

            /// <inheritdoc />
            public override long GetInt64(int ordinal) => Convert.ToInt64(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

            /// <inheritdoc />
            public override string GetName(int ordinal) => _columnNames[ordinal];

            /// <inheritdoc />
            public override int GetOrdinal(string name) {
                if (_ordinals.TryGetValue(name, out int ordinal)) {
                    return ordinal;
                }

                throw new IndexOutOfRangeException(name);
            }

            /// <inheritdoc />
            public override string GetString(int ordinal) => Convert.ToString(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty;

            /// <inheritdoc />
            public override object GetValue(int ordinal) {
                EnsureOpenRow();
                object? value = _currentRow![ordinal];
                return value == null || value == DBNull.Value ? DBNull.Value : value;
            }

            /// <inheritdoc />
            public override int GetValues(object[] values) {
                EnsureOpenRow();
                int count = Math.Min(values.Length, _fieldCount);
                for (int i = 0; i < count; i++) {
                    object? value = _currentRow![i];
                    values[i] = value == null || value == DBNull.Value ? DBNull.Value : value;
                }

                return count;
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

                if (_prefetchedIndex < _prefetchedRows.Count) {
                    _currentRow = _prefetchedRows[_prefetchedIndex++];
                    return true;
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
                _chunks.Dispose();
            }

            /// <inheritdoc />
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
                if (_closed || _nextRow > _lastRow) {
                    return false;
                }

                _ct.ThrowIfCancellationRequested();
                EnsureCurrentChunk();
                if (_currentChunk == null || _nextRow < _currentChunk.StartRow) {
                    row = CreateBlankRow(_fieldCount);
                    _nextRow++;
                    return true;
                }

                int offset = _nextRow - _currentChunk.StartRow;
                if ((uint)offset < (uint)_currentChunk.RowCount) {
                    row = NormalizeRow(_currentChunk.Rows[offset], _fieldCount);
                    _nextRow++;
                    return true;
                }

                row = CreateBlankRow(_fieldCount);
                _nextRow++;
                return true;
            }

            private void EnsureCurrentChunk() {
                while (true) {
                    if (_currentChunk != null && _nextRow < _currentChunk.StartRow + _currentChunk.RowCount) {
                        return;
                    }

                    if (!_chunks.MoveNext()) {
                        _currentChunk = null;
                        return;
                    }

                    RangeChunk chunk = _chunks.Current;
                    if (chunk.RowCount <= 0 || chunk.StartRow + chunk.RowCount <= _nextRow) {
                        continue;
                    }

                    _currentChunk = chunk;
                    return;
                }
            }

            private object GetNonDbNullValue(int ordinal) {
                object value = GetValue(ordinal);
                if (value == DBNull.Value) {
                    throw new InvalidCastException($"Column '{GetName(ordinal)}' contains DBNull.");
                }

                return value;
            }

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

            private static Type[] InferColumnTypes(IReadOnlyList<object?[]> rows, int fieldCount) {
                var types = new Type[fieldCount];
                for (int c = 0; c < fieldCount; c++) {
                    Type? inferred = null;
                    for (int r = 0; r < rows.Count; r++) {
                        inferred = MergeDataTableColumnType(inferred, rows[r][c]);
                    }

                    types[c] = inferred ?? typeof(object);
                }

                return types;
            }

            private static object?[] CreateBlankRow(int fieldCount) => new object?[fieldCount];

            private static object?[] NormalizeRow(object?[] source, int fieldCount) {
                if (source.Length == fieldCount) {
                    return source;
                }

                var values = new object?[fieldCount];
                Array.Copy(source, values, Math.Min(source.Length, values.Length));
                return values;
            }
        }
    }
}
