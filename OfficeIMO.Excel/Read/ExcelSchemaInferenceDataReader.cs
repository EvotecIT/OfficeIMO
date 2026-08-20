#nullable enable

using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Adds bounded schema inference and sampled-row replay to a forward-only Excel reader.
    /// </summary>
    internal sealed class ExcelSchemaInferenceDataReader : DbDataReader {
        private readonly DbDataReader _inner;
        private readonly List<object[]> _sampledRows;
        private readonly Type[] _columnTypes;
        private readonly CultureInfo _culture;
        private int _sampleIndex;
        private object[]? _sampledCurrentRow;
        private bool _closed;
        private bool _disposed;

        private ExcelSchemaInferenceDataReader(
            DbDataReader inner,
            int schemaSampleRows,
            int maximumSchemaSampleRows,
            long maximumBufferedCells,
            CultureInfo culture,
            CancellationToken cancellationToken) {
            _inner = inner;
            _culture = culture;

            int fieldCount = inner.FieldCount;
            if (maximumSchemaSampleRows < 0 || maximumBufferedCells <= 0L) {
                throw new InvalidOperationException("Excel data-reader safety limits must be positive.");
            }
            if (schemaSampleRows > maximumSchemaSampleRows) {
                throw new ArgumentOutOfRangeException(
                    nameof(schemaSampleRows),
                    $"Schema sample row count exceeds {maximumSchemaSampleRows}.");
            }

            _sampledRows = new List<object[]>(Math.Min(schemaSampleRows, 1024));
            while (_sampledRows.Count < schemaSampleRows && inner.Read()) {
                cancellationToken.ThrowIfCancellationRequested();
                long bufferedCells = checked((long)(_sampledRows.Count + 1) * fieldCount);
                if (bufferedCells > maximumBufferedCells) {
                    throw new InvalidDataException(
                        $"Range data-reader buffering exceeds {nameof(ExcelReadOptions.MaxDataReaderBufferedCells)}.");
                }

                var values = new object[fieldCount];
                inner.GetValues(values);
                _sampledRows.Add(values);
            }

            _columnTypes = InferColumnTypes(_sampledRows, fieldCount);
        }

        internal static DbDataReader Create(
            DbDataReader inner,
            int schemaSampleRows,
            int maximumSchemaSampleRows,
            long maximumBufferedCells,
            CultureInfo culture,
            CancellationToken cancellationToken) {
            try {
                return new ExcelSchemaInferenceDataReader(
                    inner,
                    schemaSampleRows,
                    maximumSchemaSampleRows,
                    maximumBufferedCells,
                    culture,
                    cancellationToken);
            } catch {
                inner.Dispose();
                throw;
            }
        }

        public override object this[int ordinal] => GetValue(ordinal);

        public override object this[string name] => GetValue(GetOrdinal(name));

        public override int Depth => 0;

        public override int FieldCount => _inner.FieldCount;

        public override bool HasRows => !_closed && (_sampledRows.Count > 0 || _inner.HasRows);

        public override bool IsClosed => _closed;

        public override int RecordsAffected => -1;

        public override bool GetBoolean(int ordinal) => IsSampledRow
            ? Convert.ToBoolean(GetNonDbNullValue(ordinal), _culture)
            : _inner.GetBoolean(ordinal);

        public override byte GetByte(int ordinal) => IsSampledRow
            ? Convert.ToByte(GetNonDbNullValue(ordinal), _culture)
            : _inner.GetByte(ordinal);

        public override long GetBytes(
            int ordinal,
            long dataOffset,
            byte[]? buffer,
            int bufferOffset,
            int length) =>
            throw new NotSupportedException("Excel range fields are exposed as scalar values.");

        public override char GetChar(int ordinal) => IsSampledRow
            ? Convert.ToChar(GetNonDbNullValue(ordinal), _culture)
            : _inner.GetChar(ordinal);

        public override long GetChars(
            int ordinal,
            long dataOffset,
            char[]? buffer,
            int bufferOffset,
            int length) {
            string value = Convert.ToString(GetValue(ordinal), _culture) ?? string.Empty;
            if (buffer == null) {
                return value.Length;
            }
            if (dataOffset >= value.Length || length == 0) {
                return 0;
            }

            int count = Math.Min(length, value.Length - checked((int)dataOffset));
            value.CopyTo(checked((int)dataOffset), buffer, bufferOffset, count);
            return count;
        }

        public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;

        public override DateTime GetDateTime(int ordinal) => IsSampledRow
            ? Convert.ToDateTime(GetNonDbNullValue(ordinal), _culture)
            : _inner.GetDateTime(ordinal);

        public override decimal GetDecimal(int ordinal) => IsSampledRow
            ? Convert.ToDecimal(GetNonDbNullValue(ordinal), _culture)
            : _inner.GetDecimal(ordinal);

        public override double GetDouble(int ordinal) => IsSampledRow
            ? Convert.ToDouble(GetNonDbNullValue(ordinal), _culture)
            : _inner.GetDouble(ordinal);

        [UnconditionalSuppressMessage("Trimming", "IL2063", Justification = "Inferred Excel column types are closed scalar conversion tokens; OfficeIMO never activates or reflects over their public members.")]
        [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
        public override Type GetFieldType(int ordinal) => _columnTypes[ordinal];

        public override float GetFloat(int ordinal) => IsSampledRow
            ? Convert.ToSingle(GetNonDbNullValue(ordinal), _culture)
            : _inner.GetFloat(ordinal);

        public override Guid GetGuid(int ordinal) {
            if (!IsSampledRow) {
                return _inner.GetGuid(ordinal);
            }

            object value = GetNonDbNullValue(ordinal);
            return value is Guid guid ? guid : Guid.Parse(Convert.ToString(value, _culture)!);
        }

        public override short GetInt16(int ordinal) => IsSampledRow
            ? Convert.ToInt16(GetNonDbNullValue(ordinal), _culture)
            : _inner.GetInt16(ordinal);

        public override int GetInt32(int ordinal) => IsSampledRow
            ? Convert.ToInt32(GetNonDbNullValue(ordinal), _culture)
            : _inner.GetInt32(ordinal);

        public override long GetInt64(int ordinal) => IsSampledRow
            ? Convert.ToInt64(GetNonDbNullValue(ordinal), _culture)
            : _inner.GetInt64(ordinal);

        public override string GetName(int ordinal) => _inner.GetName(ordinal);

        public override int GetOrdinal(string name) => _inner.GetOrdinal(name);

        public override string GetString(int ordinal) => IsSampledRow
            ? Convert.ToString(GetNonDbNullValue(ordinal), _culture) ?? string.Empty
            : _inner.GetString(ordinal);

        public override object GetValue(int ordinal) => IsSampledRow
            ? _sampledCurrentRow![ordinal]
            : _inner.GetValue(ordinal);

        public override int GetValues(object[] values) {
            if (!IsSampledRow) {
                return _inner.GetValues(values);
            }

            int count = Math.Min(values.Length, FieldCount);
            Array.Copy(_sampledCurrentRow!, values, count);
            return count;
        }

        public override bool IsDBNull(int ordinal) {
            object value = GetValue(ordinal);
            return value == null || ReferenceEquals(value, DBNull.Value);
        }

        public override bool NextResult() => false;

        public override bool Read() {
            if (_closed) {
                return false;
            }
            if (_sampleIndex < _sampledRows.Count) {
                _sampledCurrentRow = _sampledRows[_sampleIndex++];
                return true;
            }

            _sampledCurrentRow = null;
            return _inner.Read();
        }

        public override void Close() {
            if (_closed) {
                return;
            }

            _closed = true;
            _sampledCurrentRow = null;
            _inner.Close();
        }

        [UnconditionalSuppressMessage("Trimming", "IL2111", Justification = "The schema table stores Type values as data and does not reflect over Type.TypeInitializer or other Type members.")]
        public override DataTable GetSchemaTable() =>
            ExcelDataReaderSchemaTable.Create(FieldCount, GetName, GetFieldType);

        public override IEnumerator GetEnumerator() {
            while (Read()) {
                yield return this;
            }
        }

        protected override void Dispose(bool disposing) {
            if (disposing && !_disposed) {
                _disposed = true;
                Close();
                _inner.Dispose();
            }

            base.Dispose(disposing);
        }

        private bool IsSampledRow => _sampledCurrentRow != null;

        private object GetNonDbNullValue(int ordinal) {
            object value = GetValue(ordinal);
            if (value == null || ReferenceEquals(value, DBNull.Value)) {
                throw new InvalidCastException($"Column '{GetName(ordinal)}' contains DBNull.");
            }

            return value;
        }

        private static Type[] InferColumnTypes(IReadOnlyList<object[]> rows, int fieldCount) {
            var types = new Type[fieldCount];
            for (int column = 0; column < fieldCount; column++) {
                Type? inferred = null;
                for (int row = 0; row < rows.Count; row++) {
                    inferred = ExcelSheetReader.MergeDataTableColumnType(inferred, rows[row][column]);
                    if (inferred == typeof(object)) {
                        break;
                    }
                }

                types[column] = inferred ?? typeof(object);
            }

            return types;
        }
    }
}
