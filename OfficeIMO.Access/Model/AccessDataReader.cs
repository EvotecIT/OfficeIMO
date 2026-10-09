using System.Collections;
using System.Data;
using System.Data.Common;
using System.Globalization;

namespace OfficeIMO.Access {
    /// <summary>Forward-only typed reader for modeled or qualified native rows. Disposal releases its document lease.</summary>
    public sealed class AccessDataReader : DbDataReader {
        private readonly AccessTable _table;
        private readonly CancellationToken _cancellation;
        private readonly IAccessRowCursor _cursor;
        private bool _closed;
        internal AccessDataReader(AccessTable table, CancellationToken cancellation, IAccessRowCursor? cursor = null) {
            _table = table; _cancellation = cancellation; table.Document.AcquireReader();
            try { AccessNativeTable? source = table.Document.ResolveNativeReadTable(table.NativeTable, cancellation); _cursor = cursor ?? (source == null ? new AccessModeledRowCursor(table) : new AccessNativeRowCursor(source, cancellation)); }
            catch { table.Document.ReleaseReader(); throw; }
        }
        private void Check() {
            if (_closed) throw new ObjectDisposedException(nameof(AccessDataReader));
            _table.EnsureAttached(); _cancellation.ThrowIfCancellationRequested();
        }
        private AccessColumn Column(int ordinal) { Check(); return _table.Columns[ordinal]; }
        /// <summary>Distinguishes an omitted input value from an explicit null.</summary>
        public bool IsSpecified(int ordinal) { Column(ordinal); return _cursor.IsSpecified(ordinal); }
        /// <inheritdoc />
        public override bool Read() { Check(); return _cursor.Read(_cancellation); }
        /// <inheritdoc />
        public override Task<bool> ReadAsync(CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested(); Check();
            using CancellationTokenSource linked = CancellationTokenSource.CreateLinkedTokenSource(_cancellation, cancellationToken);
            return Task.FromResult(_cursor.Read(linked.Token));
        }
        /// <inheritdoc />
        public override int FieldCount { get { Check(); return _table.Columns.Count; } }
        /// <inheritdoc />
        public override bool HasRows { get { Check(); return _cursor.HasRows; } }
        /// <inheritdoc />
        public override bool IsClosed => _closed;
        /// <inheritdoc />
        public override int Depth => 0;
        /// <inheritdoc />
        public override int RecordsAffected => -1;
        /// <inheritdoc />
        public override object this[int ordinal] => GetValue(ordinal);
        /// <inheritdoc />
        public override object this[string name] => GetValue(GetOrdinal(name));
        /// <inheritdoc />
        public override string GetName(int ordinal) => Column(ordinal).Name;
        /// <inheritdoc />
        public override int GetOrdinal(string name) {
            Check(); for (int i = 0; i < FieldCount; i++) if (StringComparer.OrdinalIgnoreCase.Equals(GetName(i), name)) return i;
            throw new IndexOutOfRangeException($"No field named '{name}'.");
        }
        /// <inheritdoc />
        public override object GetValue(int ordinal) { Column(ordinal); return AccessTable.CopyValue(_cursor.GetValue(ordinal, _cancellation)) ?? DBNull.Value; }
        /// <inheritdoc />
        public override int GetValues(object[] values) {
            if (values == null) throw new ArgumentNullException(nameof(values));
            int count = Math.Min(values.Length, FieldCount); for (int i = 0; i < count; i++) values[i] = GetValue(i); return count;
        }
        /// <inheritdoc />
        public override bool IsDBNull(int ordinal) { Column(ordinal); return _cursor.IsNull(ordinal); }
        /// <inheritdoc />
        public override Type GetFieldType(int ordinal) => Column(ordinal).DataType switch {
            AccessDataType.AutoNumber or AccessDataType.Int32 => typeof(int),
            AccessDataType.ShortText or AccessDataType.LongText => typeof(string), AccessDataType.Currency => typeof(decimal),
            AccessDataType.Double => typeof(double), AccessDataType.Boolean => typeof(bool), AccessDataType.DateTime => typeof(DateTime),
            AccessDataType.Guid => typeof(Guid), AccessDataType.Binary => typeof(byte[]), AccessDataType.Byte => typeof(byte),
            AccessDataType.Int16 => typeof(short), AccessDataType.Int64 => typeof(long), AccessDataType.Single => typeof(float),
            AccessDataType.Decimal => Column(ordinal).Precision < 1 || Column(ordinal).Precision > 28 || Column(ordinal).Scale > 28 ? typeof(AccessOpaqueValue) : typeof(decimal), AccessDataType.ExtendedDateTime => typeof(DateTime),
            AccessDataType.Complex => Column(ordinal).ComplexDefinition == null ? typeof(AccessOpaqueValue) : typeof(AccessComplexValue), AccessDataType.Unknown => typeof(AccessOpaqueValue), _ => throw new NotSupportedException()
        };
        /// <inheritdoc />
        public override string GetDataTypeName(int ordinal) => Column(ordinal).DataType.ToString();
        /// <inheritdoc />
        public override bool NextResult() { Check(); return false; }
        /// <inheritdoc />
        public override IEnumerator GetEnumerator() => new DbEnumerator(this, false);
        /// <inheritdoc />
        public override DataTable? GetSchemaTable() { Check(); return null; }
        /// <inheritdoc />
        public override bool GetBoolean(int ordinal) => (bool)GetValue(ordinal);
        /// <inheritdoc />
        public override byte GetByte(int ordinal) => (byte)GetValue(ordinal);
        /// <inheritdoc />
        public override char GetChar(int ordinal) => (char)GetValue(ordinal);
        /// <inheritdoc />
        public override DateTime GetDateTime(int ordinal) => (DateTime)GetValue(ordinal);
        /// <inheritdoc />
        public override decimal GetDecimal(int ordinal) => (decimal)GetValue(ordinal);
        /// <inheritdoc />
        public override double GetDouble(int ordinal) => (double)GetValue(ordinal);
        /// <inheritdoc />
        public override float GetFloat(int ordinal) => (float)GetValue(ordinal);
        /// <inheritdoc />
        public override Guid GetGuid(int ordinal) => (Guid)GetValue(ordinal);
        /// <inheritdoc />
        public override short GetInt16(int ordinal) => (short)GetValue(ordinal);
        /// <inheritdoc />
        public override int GetInt32(int ordinal) => (int)GetValue(ordinal);
        /// <inheritdoc />
        public override long GetInt64(int ordinal) => (long)GetValue(ordinal);
        /// <inheritdoc />
        public override string GetString(int ordinal) => (string)GetValue(ordinal);
        /// <inheritdoc />
        public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) {
            using Stream stream = GetStream(ordinal);
            if (dataOffset < 0 || dataOffset > stream.Length) throw new ArgumentOutOfRangeException(nameof(dataOffset));
            if (buffer == null) return stream.Length;
            if (bufferOffset < 0 || length < 0 || bufferOffset > buffer.Length - length) throw new ArgumentOutOfRangeException(nameof(bufferOffset));
            byte[] skip = new byte[checked((int)Math.Min(dataOffset, 8192))]; long remaining = dataOffset;
            while (remaining > 0) { int read = stream.Read(skip, 0, (int)Math.Min(skip.Length, remaining)); if (read == 0) throw new InvalidDataException("Native Access binary value is truncated."); remaining -= read; }
            return stream.Read(buffer, bufferOffset, length);
        }
        /// <inheritdoc />
        public override Stream GetStream(int ordinal) { Column(ordinal); return _cursor.OpenBinary(ordinal, _cancellation); }
        /// <inheritdoc />
        public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) => CopyRange(GetString(ordinal).ToCharArray(), dataOffset, buffer, bufferOffset, length);
        private static long CopyRange<T>(T[] source, long dataOffset, T[]? buffer, int bufferOffset, int length) {
            if (dataOffset < 0 || dataOffset > source.LongLength) throw new ArgumentOutOfRangeException(nameof(dataOffset));
            if (buffer == null) return source.LongLength;
            if (bufferOffset < 0 || length < 0 || bufferOffset > buffer.Length - length) throw new ArgumentOutOfRangeException(nameof(bufferOffset));
            int count = (int)Math.Min(source.LongLength - dataOffset, length);
            Array.Copy(source, dataOffset, buffer, bufferOffset, count); return count;
        }
        /// <inheritdoc />
        public override void Close() { if (_closed) return; _closed = true; try { _cursor.Dispose(); } finally { _table.Document.ReleaseReader(); } }
        /// <inheritdoc />
        protected override void Dispose(bool disposing) { if (disposing) Close(); base.Dispose(disposing); }
    }
}
