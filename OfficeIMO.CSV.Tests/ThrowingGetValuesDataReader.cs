using System;
using System.Collections;
using System.Data;
using System.Data.Common;

namespace OfficeIMO.CSV.Tests;

internal sealed class ThrowingGetValuesDataReader : DbDataReader
{
    private readonly string[] _headers;
    private readonly object?[][] _rows;
    private readonly Type[] _fieldTypes;
    private readonly Action<int>? _afterRead;
    private readonly bool _throwOnGetDataTypeName;
    private int _rowIndex = -1;
    private bool _closed;

    public int GetValueCallCount { get; private set; }

    public ThrowingGetValuesDataReader(
        string[] headers,
        object?[][] rows,
        Action<int>? afterRead = null,
        bool throwOnGetDataTypeName = false)
    {
        _headers = headers ?? throw new ArgumentNullException(nameof(headers));
        _rows = rows ?? throw new ArgumentNullException(nameof(rows));
        _fieldTypes = CreateFieldTypes(headers, rows);
        _afterRead = afterRead;
        _throwOnGetDataTypeName = throwOnGetDataTypeName;
    }

    public override object this[int ordinal] => GetValue(ordinal);

    public override object this[string name] => GetValue(GetOrdinal(name));

    public override int Depth => 0;

    public override int FieldCount => _headers.Length;

    public override bool HasRows => _rows.Length > 0;

    public override bool IsClosed => _closed;

    public override int RecordsAffected => -1;

    public override bool GetBoolean(int ordinal) => GetRequiredValue<bool>(ordinal);

    public override byte GetByte(int ordinal) => GetRequiredValue<byte>(ordinal);

    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) => throw new NotSupportedException();

    public override char GetChar(int ordinal) => (char)GetValue(ordinal);

    public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) => throw new NotSupportedException();

    public override string GetDataTypeName(int ordinal) => _throwOnGetDataTypeName
        ? throw new NotImplementedException()
        : GetFieldType(ordinal).Name;

    public override DateTime GetDateTime(int ordinal) => GetRequiredValue<DateTime>(ordinal);

    public override decimal GetDecimal(int ordinal) => GetRequiredValue<decimal>(ordinal);

    public override double GetDouble(int ordinal) => GetRequiredValue<double>(ordinal);

    public override IEnumerator GetEnumerator()
    {
        while (Read())
        {
            yield return this;
        }
    }

    public override Type GetFieldType(int ordinal) => _fieldTypes[ordinal];

    public override float GetFloat(int ordinal) => GetRequiredValue<float>(ordinal);

    public override Guid GetGuid(int ordinal) => GetRequiredValue<Guid>(ordinal);

    public override short GetInt16(int ordinal) => GetRequiredValue<short>(ordinal);

    public override int GetInt32(int ordinal) => GetRequiredValue<int>(ordinal);

    public override long GetInt64(int ordinal) => GetRequiredValue<long>(ordinal);

    public override string GetName(int ordinal) => _headers[ordinal];

    public override int GetOrdinal(string name)
    {
        for (var i = 0; i < _headers.Length; i++)
        {
            if (string.Equals(_headers[i], name, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        throw new IndexOutOfRangeException(name);
    }

    public override string GetString(int ordinal) => GetRequiredValue<string>(ordinal);

    public override object GetValue(int ordinal)
    {
        GetValueCallCount++;
        var value = CurrentRow[ordinal];
        return value ?? DBNull.Value;
    }

    public override int GetValues(object[] values) => throw new NotSupportedException("GetValues should not be required for CSV data reader export.");

    public override bool IsDBNull(int ordinal)
    {
        var value = CurrentRow[ordinal];
        return value is null || ReferenceEquals(value, DBNull.Value);
    }

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
        _afterRead?.Invoke(_rowIndex);
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

    private T GetRequiredValue<T>(int ordinal)
    {
        var value = CurrentRow[ordinal];
        if (value is null || ReferenceEquals(value, DBNull.Value))
        {
            throw new InvalidCastException("The requested field contains a database null.");
        }

        return (T)value;
    }

    private static Type[] CreateFieldTypes(string[] headers, object?[][] rows)
    {
        var types = new Type[headers.Length];
        if (rows.Length == 0)
        {
            for (var i = 0; i < types.Length; i++)
            {
                types[i] = typeof(string);
            }

            return types;
        }

        var firstRow = rows[0];
        for (var i = 0; i < headers.Length; i++)
        {
            var value = firstRow[i];
            types[i] = value is null || ReferenceEquals(value, DBNull.Value)
                ? typeof(string)
                : value.GetType();
        }

        return types;
    }
}
