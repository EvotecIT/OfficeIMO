using System.Collections;
using System.Data;
using System.Data.Common;
using System.Globalization;

namespace OfficeIMO.Access;

/// <summary>Forward-only modeled-row reader. Disposal releases an edit-blocking lease; this slice does not decode native rows.</summary>
public sealed class AccessDataReader : DbDataReader {
    private readonly AccessTable _table;
    private readonly CancellationToken _cancellation;
    private int _row = -1;
    private bool _closed;
    internal AccessDataReader(AccessTable table, CancellationToken cancellation) { _table = table; _cancellation = cancellation; table.Document.AcquireReader(); }
    private void Check() {
        if (_closed) throw new ObjectDisposedException(nameof(AccessDataReader));
        _table.EnsureAttached(); _cancellation.ThrowIfCancellationRequested();
    }
    private AccessColumn Column(int ordinal) { Check(); return _table.Columns[ordinal]; }
    private Dictionary<string, object?> Current {
        get { Check(); if (_row < 0 || _row >= _table.Rows.Count) throw new InvalidOperationException("Call Read before accessing a current row."); return _table.Rows[_row]; }
    }
    /// <summary>Distinguishes an omitted input value from an explicit null.</summary>
    public bool IsSpecified(int ordinal) => Current.ContainsKey(Column(ordinal).Name);
    /// <inheritdoc />
    public override bool Read() { Check(); if (_row < _table.Rows.Count) _row++; return _row < _table.Rows.Count; }
    /// <inheritdoc />
    public override Task<bool> ReadAsync(CancellationToken cancellationToken) { cancellationToken.ThrowIfCancellationRequested(); return Task.FromResult(Read()); }
    /// <inheritdoc />
    public override int FieldCount { get { Check(); return _table.Columns.Count; } }
    /// <inheritdoc />
    public override bool HasRows { get { Check(); return _table.Rows.Count != 0; } }
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
    public override object GetValue(int ordinal) => Current.TryGetValue(Column(ordinal).Name, out object? value) ? AccessTable.CopyValue(value) ?? DBNull.Value : DBNull.Value;
    /// <inheritdoc />
    public override int GetValues(object[] values) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        int count = Math.Min(values.Length, FieldCount); for (int i = 0; i < count; i++) values[i] = GetValue(i); return count;
    }
    /// <inheritdoc />
    public override bool IsDBNull(int ordinal) => GetValue(ordinal) == DBNull.Value;
    /// <inheritdoc />
    public override Type GetFieldType(int ordinal) => Column(ordinal).DataType switch {
        AccessDataType.AutoNumber or AccessDataType.Int32 => typeof(int),
        AccessDataType.ShortText or AccessDataType.LongText => typeof(string), AccessDataType.Currency => typeof(decimal),
        AccessDataType.Double => typeof(double), AccessDataType.Boolean => typeof(bool), AccessDataType.DateTime => typeof(DateTime),
        AccessDataType.Guid => typeof(Guid), AccessDataType.Binary => typeof(byte[]), _ => throw new NotSupportedException()
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
    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) => CopyRange((byte[])GetValue(ordinal), dataOffset, buffer, bufferOffset, length);
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
    public override void Close() { if (_closed) return; _closed = true; _table.Document.ReleaseReader(); }
    /// <inheritdoc />
    protected override void Dispose(bool disposing) { if (disposing) Close(); base.Dispose(disposing); }
}
