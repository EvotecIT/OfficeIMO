#nullable enable

using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
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
            if (chunkRows <= 0 || chunkRows > _opt.MaxDataReaderChunkRows) {
                throw new ArgumentOutOfRangeException(nameof(chunkRows),
                    $"Chunk row count must be between 1 and {_opt.MaxDataReaderChunkRows}.");
            }
            if (schemaSampleRows < 0) {
                throw new ArgumentOutOfRangeException(nameof(schemaSampleRows), "Schema sample row count cannot be negative.");
            }
            if (schemaSampleRows > _opt.MaxDataReaderSchemaSampleRows) {
                throw new ArgumentOutOfRangeException(nameof(schemaSampleRows),
                    $"Schema sample row count exceeds {_opt.MaxDataReaderSchemaSampleRows}.");
            }
            if (_opt.MaxDataReaderColumns <= 0 || _opt.MaxDataReaderChunkRows <= 0 ||
                _opt.MaxDataReaderSchemaSampleRows < 0 || _opt.MaxDataReaderBufferedCells <= 0L) {
                throw new InvalidOperationException("Excel data-reader safety limits must be positive.");
            }

            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) {
                throw new ArgumentException($"Invalid range '{a1Range}'.", nameof(a1Range));
            }

            int cols = c2 - c1 + 1;
            int rows = r2 - r1 + 1;
            if (cols > _opt.MaxDataReaderColumns) {
                throw new InvalidDataException($"Range data-reader column count exceeds {nameof(ExcelReadOptions.MaxDataReaderColumns)}.");
            }
            long sampleCells = (long)Math.Min(Math.Max(0, rows - (headersInFirstRow ? 1 : 0)), schemaSampleRows) * cols;
            if (sampleCells > _opt.MaxDataReaderBufferedCells) {
                throw new InvalidDataException($"Range data-reader buffering exceeds {nameof(ExcelReadOptions.MaxDataReaderBufferedCells)}.");
            }
            if (schemaSampleRows == 0
                && rows > BufferedRangeStreamRowLimit
                && mode != OfficeIMO.Excel.ExecutionMode.Parallel
                && CanUseRangeStreamXmlReader()) {
                if (cols > _opt.MaxDataReaderBufferedCells) {
                    throw new InvalidDataException($"Range data-reader buffering exceeds {nameof(ExcelReadOptions.MaxDataReaderBufferedCells)}.");
                }

                return new ExcelXmlRangeDataReader(this, r1, c1, r2, c2, cols, headersInFirstRow, _opt, ct);
            }

            long chunkCells = (long)Math.Min(rows, chunkRows) * cols;
            if (chunkCells > _opt.MaxDataReaderBufferedCells) {
                throw new InvalidDataException($"Range data-reader buffering exceeds {nameof(ExcelReadOptions.MaxDataReaderBufferedCells)}.");
            }

            IEnumerable<RangeChunk> chunks = ReadRangeStreamForDataReader(a1Range, chunkRows, mode, ct);
            return new ExcelRangeDataReader(chunks, r1, r2, cols, headersInFirstRow, schemaSampleRows, _opt, ct);
        }

        private IEnumerable<RangeChunk> ReadRangeStreamForDataReader(
            string a1Range,
            int chunkRows,
            OfficeIMO.Excel.ExecutionMode? mode,
            CancellationToken ct) {
            if (chunkRows <= 0) {
                throw new ArgumentOutOfRangeException(nameof(chunkRows), "Chunk row count must be greater than zero.");
            }

            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) {
                yield break;
            }

            if (ct.CanBeCanceled) {
                ct.ThrowIfCancellationRequested();
            }

            if (mode != OfficeIMO.Excel.ExecutionMode.Parallel
                && CanUseRangeStreamXmlReader()
                && RowsAreSortedWithinRangeXmlFast(r1, r2, ct)) {
                foreach (var chunk in ReadRangeStreamXmlFast(r1, c1, r2, c2, chunkRows, ct)) {
                    yield return chunk;
                }

                yield break;
            }

            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData is null) {
                yield break;
            }

            if (mode != OfficeIMO.Excel.ExecutionMode.Parallel
                && RowsAreSortedWithinRange(sheetData, r1, r2, ct)) {
                foreach (var chunk in ReadSortedDomRangeStream(sheetData, r1, c1, r2, c2, chunkRows, ct)) {
                    yield return chunk;
                }

                yield break;
            }

            foreach (var chunk in ReadRangeStream(a1Range, chunkRows, OfficeIMO.Excel.ExecutionMode.Sequential, ct)) {
                yield return chunk;
            }
        }

        private static int ConvertDataReaderInt32(object value) {
            if (value is int intValue) {
                return intValue;
            }

            if (value is double doubleValue && doubleValue >= int.MinValue && doubleValue <= int.MaxValue) {
                return ConvertDataReaderInt32(doubleValue);
            }

            return Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }

        private static int ConvertDataReaderInt32(double value) {
            if (value >= int.MinValue && value <= int.MaxValue) {
                int candidate = (int)value;
                if (candidate == value) {
                    return candidate;
                }
            }

            return Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }

        private static Dictionary<string, int> CreateOrdinalMap(string[] columnNames) {
            var ordinals = new Dictionary<string, int>(columnNames.Length, StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < columnNames.Length; i++) {
                ordinals[columnNames[i]] = i;
            }

            return ordinals;
        }

        private IEnumerable<RangeChunk> ReadSortedDomRangeStream(
            SheetData sheetData,
            int r1,
            int c1,
            int r2,
            int c2,
            int chunkRows,
            CancellationToken ct) {
            int currentWindow = -1;
            var rows = new List<Row>();

            foreach (var row in sheetData.Elements<Row>()) {
                if (ct.CanBeCanceled) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    break;
                }

                int window = (rowIndex - r1) / chunkRows;
                if (currentWindow >= 0 && window != currentWindow) {
                    yield return ConvertSortedDomChunk(rows, currentWindow, r1, c1, r2, c2, chunkRows, ct);
                    rows.Clear();
                }

                currentWindow = window;
                rows.Add(row);
            }

            if (rows.Count > 0) {
                yield return ConvertSortedDomChunk(rows, currentWindow, r1, c1, r2, c2, chunkRows, ct);
            }
        }

        private RangeChunk ConvertSortedDomChunk(
            IReadOnlyList<Row> rows,
            int windowIndex,
            int r1,
            int c1,
            int r2,
            int c2,
            int chunkRows,
            CancellationToken ct) {
            int startRow = r1 + (windowIndex * chunkRows);
            int endRow = Math.Min(startRow + chunkRows - 1, r2);
            int height = endRow - startRow + 1;
            int width = c2 - c1 + 1;
            if (height <= 0 || width <= 0) {
                return new RangeChunk(startRow, 0, c1, width, Array.Empty<object?[]>());
            }

            var values = new object?[height][];
            for (int i = 0; i < height; i++) {
                values[i] = new object?[width];
            }

            foreach (var row in rows) {
                if (ct.CanBeCanceled) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                int rowOffset = rowIndex - startRow;
                if ((uint)rowOffset >= (uint)height) {
                    continue;
                }

                foreach (var cell in row.Elements<Cell>()) {
                    int column = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (column < c1 || column > c2) {
                        continue;
                    }

                    if (TryConvertCell(cell, out object? value)) {
                        values[rowOffset][column - c1] = value ?? values[rowOffset][column - c1];
                    }
                }
            }

            return new RangeChunk(startRow, height, c1, width, values);
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
            private readonly List<object?[]> _prefetchedRows;
            private readonly object?[] _blankRow;
            private Dictionary<string, int>? _ordinals;
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
                _blankRow = new object?[fieldCount];

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
            public override bool GetBoolean(int ordinal) {
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
                object value = GetNonDbNullValue(ordinal);
                return value is DateTime dateTime ? dateTime : Convert.ToDateTime(value, CultureInfo.InvariantCulture);
            }

            /// <inheritdoc />
            public override decimal GetDecimal(int ordinal) => Convert.ToDecimal(GetNonDbNullValue(ordinal), CultureInfo.InvariantCulture);

            /// <inheritdoc />
            public override double GetDouble(int ordinal) {
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
                object? value = _currentRow![ordinal];
                return ToDataReaderValue(value);
            }

            /// <inheritdoc />
            public override int GetValues(object[] values) {
                EnsureOpenRow();
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
                if (_closed || _nextRow > _lastRow) {
                    return false;
                }

                _ct.ThrowIfCancellationRequested();
                EnsureCurrentChunk();
                if (_currentChunk == null || _nextRow < _currentChunk.StartRow) {
                    row = _blankRow;
                    _nextRow++;
                    return true;
                }

                int offset = _nextRow - _currentChunk.StartRow;
                if ((uint)offset < (uint)_currentChunk.RowCount) {
                    row = NormalizeRow(_currentChunk.Rows[offset], _fieldCount);
                    _nextRow++;
                    return true;
                }

                row = _blankRow;
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
                EnsureOpenRow();
                object? value = _currentRow![ordinal];
                if (value == null || value == DBNull.Value) {
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

            private static object?[] NormalizeRow(object?[] source, int fieldCount) {
                if (source.Length == fieldCount) {
                    return source;
                }

                var values = new object?[fieldCount];
                Array.Copy(source, values, Math.Min(source.Length, values.Length));
                return values;
            }
        }

        private static object ToDataReaderValue(object? value)
            => value == null || ReferenceEquals(value, DBNull.Value) ? DBNull.Value : value;

        private static int CopyDataReaderValues(object?[] row, int fieldCount, object[] values) {
            int count = Math.Min(values.Length, fieldCount);
            if (count == fieldCount) {
                if (fieldCount == 8) {
                    values[0] = ToDataReaderValue(row[0]);
                    values[1] = ToDataReaderValue(row[1]);
                    values[2] = ToDataReaderValue(row[2]);
                    values[3] = ToDataReaderValue(row[3]);
                    values[4] = ToDataReaderValue(row[4]);
                    values[5] = ToDataReaderValue(row[5]);
                    values[6] = ToDataReaderValue(row[6]);
                    values[7] = ToDataReaderValue(row[7]);
                    return count;
                }

                if (fieldCount == 3) {
                    values[0] = ToDataReaderValue(row[0]);
                    values[1] = ToDataReaderValue(row[1]);
                    values[2] = ToDataReaderValue(row[2]);
                    return count;
                }
            }

            for (int i = 0; i < count; i++) {
                values[i] = ToDataReaderValue(row[i]);
            }

            return count;
        }
    }
}
