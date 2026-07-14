#nullable enable

using System.Collections;
using System.Data;
using System.Data.Common;

namespace OfficeIMO.CSV.Benchmarks;

internal sealed class BenchmarkArrayDataReader : DbDataReader
{
    private readonly string[] _headers;
    private readonly object?[][] _rows;
    private readonly Type[] _fieldTypes;
    private readonly Dictionary<string, int> _ordinals;
    private int _rowIndex = -1;
    private bool _closed;

    public BenchmarkArrayDataReader(string[] headers, object?[][] rows, Type[]? fieldTypes = null)
    {
        _headers = headers;
        _rows = rows;
        if (fieldTypes != null && fieldTypes.Length != headers.Length)
        {
            throw new ArgumentException("Field type count must match the header count.", nameof(fieldTypes));
        }

        _fieldTypes = fieldTypes ?? CreateFieldTypes(headers, rows);
        _ordinals = new Dictionary<string, int>(headers.Length, StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < headers.Length; i++)
        {
            _ordinals[headers[i]] = i;
        }
    }

    public override object this[int ordinal] => GetValue(ordinal);

    public override object this[string name] => GetValue(GetOrdinal(name));

    public override int Depth => 0;

    public override int FieldCount => _headers.Length;

    public override bool HasRows => _rows.Length > 0;

    public override bool IsClosed => _closed;

    public override int RecordsAffected => -1;

    public override bool GetBoolean(int ordinal) => (bool)CurrentRow[ordinal]!;

    public override byte GetByte(int ordinal) => (byte)CurrentRow[ordinal]!;

    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length)
    {
        throw new NotSupportedException();
    }

    public override char GetChar(int ordinal) => (char)CurrentRow[ordinal]!;

    public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length)
    {
        throw new NotSupportedException();
    }

    public override string GetDataTypeName(int ordinal) => _fieldTypes[ordinal].Name;

    public override DateTime GetDateTime(int ordinal) => (DateTime)CurrentRow[ordinal]!;

    public override decimal GetDecimal(int ordinal) => (decimal)CurrentRow[ordinal]!;

    public override double GetDouble(int ordinal) => (double)CurrentRow[ordinal]!;

    public override IEnumerator GetEnumerator()
    {
        while (Read())
        {
            yield return this;
        }
    }

    public override Type GetFieldType(int ordinal) => _fieldTypes[ordinal];

    public override float GetFloat(int ordinal) => (float)CurrentRow[ordinal]!;

    public override Guid GetGuid(int ordinal) => (Guid)CurrentRow[ordinal]!;

    public override short GetInt16(int ordinal) => (short)CurrentRow[ordinal]!;

    public override int GetInt32(int ordinal) => (int)CurrentRow[ordinal]!;

    public override long GetInt64(int ordinal) => (long)CurrentRow[ordinal]!;

    public override string GetName(int ordinal) => _headers[ordinal];

    public override int GetOrdinal(string name)
    {
        if (_ordinals.TryGetValue(name, out var ordinal))
        {
            return ordinal;
        }

        throw new IndexOutOfRangeException(name);
    }

    public override string GetString(int ordinal) => (string)CurrentRow[ordinal]!;

    public override object GetValue(int ordinal) => CurrentRow[ordinal] ?? DBNull.Value;

    public override int GetValues(object[] values)
    {
        var row = CurrentRow;
        var count = Math.Min(values.Length, row.Length);
        for (var i = 0; i < count; i++)
        {
            values[i] = row[i] ?? DBNull.Value;
        }

        return count;
    }

    public override bool IsDBNull(int ordinal) => CurrentRow[ordinal] == null;

    public override bool NextResult() => false;

    public override bool Read()
    {
        if (_closed)
        {
            return false;
        }

        var next = _rowIndex + 1;
        if (next >= _rows.Length)
        {
            return false;
        }

        _rowIndex = next;
        return true;
    }

    public override void Close()
    {
        _closed = true;
    }

    public override DataTable? GetSchemaTable() => null;

    private object?[] CurrentRow
    {
        get
        {
            if (_rowIndex < 0 || _rowIndex >= _rows.Length)
            {
                throw new InvalidOperationException("The reader is not positioned on a row.");
            }

            return _rows[_rowIndex];
        }
    }

    private static Type[] CreateFieldTypes(string[] headers, object?[][] rows)
    {
        var types = new Type[headers.Length];
        if (rows.Length == 0)
        {
            Array.Fill(types, typeof(string));
            return types;
        }

        var firstRow = rows[0];
        for (var i = 0; i < headers.Length; i++)
        {
            types[i] = firstRow[i]?.GetType() ?? typeof(string);
        }

        return types;
    }
}
