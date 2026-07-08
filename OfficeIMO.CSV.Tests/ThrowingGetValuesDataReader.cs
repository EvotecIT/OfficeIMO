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
    private int _rowIndex = -1;
    private bool _closed;

    public ThrowingGetValuesDataReader(string[] headers, object?[][] rows)
    {
        _headers = headers ?? throw new ArgumentNullException(nameof(headers));
        _rows = rows ?? throw new ArgumentNullException(nameof(rows));
        _fieldTypes = CreateFieldTypes(headers, rows);
    }

    public override object this[int ordinal] => GetValue(ordinal);

    public override object this[string name] => GetValue(GetOrdinal(name));

    public override int Depth => 0;

    public override int FieldCount => _headers.Length;

    public override bool HasRows => _rows.Length > 0;

    public override bool IsClosed => _closed;

    public override int RecordsAffected => -1;

    public override bool GetBoolean(int ordinal) => (bool)GetValue(ordinal);

    public override byte GetByte(int ordinal) => (byte)GetValue(ordinal);

    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) => throw new NotSupportedException();

    public override char GetChar(int ordinal) => (char)GetValue(ordinal);

    public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) => throw new NotSupportedException();

    public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;

    public override DateTime GetDateTime(int ordinal) => (DateTime)GetValue(ordinal);

    public override decimal GetDecimal(int ordinal) => (decimal)GetValue(ordinal);

    public override double GetDouble(int ordinal) => (double)GetValue(ordinal);

    public override IEnumerator GetEnumerator()
    {
        while (Read())
        {
            yield return this;
        }
    }

    public override Type GetFieldType(int ordinal) => _fieldTypes[ordinal];

    public override float GetFloat(int ordinal) => (float)GetValue(ordinal);

    public override Guid GetGuid(int ordinal) => (Guid)GetValue(ordinal);

    public override short GetInt16(int ordinal) => (short)GetValue(ordinal);

    public override int GetInt32(int ordinal) => (int)GetValue(ordinal);

    public override long GetInt64(int ordinal) => (long)GetValue(ordinal);

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

    public override string GetString(int ordinal) => (string)GetValue(ordinal);

    public override object GetValue(int ordinal)
    {
        var value = CurrentRow[ordinal];
        return value ?? DBNull.Value;
    }

    public override int GetValues(object[] values) => throw new NotSupportedException("GetValues should not be required for CSV data reader export.");

    public override bool IsDBNull(int ordinal) => ReferenceEquals(GetValue(ordinal), DBNull.Value);

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
