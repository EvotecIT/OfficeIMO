using System;
using System.Data;

namespace PSWriteOffice.Services.Excel;

/// <summary>Closes a forward-only reader automatically when it reaches the end of its result.</summary>
internal sealed class OwnedDataReader : IDataReader
{
    private readonly IDataReader _reader;
    private bool _disposed;

    internal OwnedDataReader(IDataReader reader)
    {
        _reader = reader ?? throw new ArgumentNullException(nameof(reader));
    }

    public object this[int i] => _reader[i];
    public object this[string name] => _reader[name];
    public int Depth => _reader.Depth;
    public bool IsClosed => _disposed || _reader.IsClosed;
    public int RecordsAffected => _reader.RecordsAffected;
    public int FieldCount => _reader.FieldCount;
    public void Close() => Dispose();

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _reader.Dispose();
    }

    public DataTable? GetSchemaTable() => _reader.GetSchemaTable();
    public bool NextResult() => _reader.NextResult();

    public bool Read()
    {
        if (_disposed)
        {
            return false;
        }

        if (_reader.Read())
        {
            return true;
        }

        Dispose();
        return false;
    }

    public bool GetBoolean(int i) => _reader.GetBoolean(i);
    public byte GetByte(int i) => _reader.GetByte(i);
    public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferOffset, int length) =>
        _reader.GetBytes(i, fieldOffset, buffer, bufferOffset, length);
    public char GetChar(int i) => _reader.GetChar(i);
    public long GetChars(int i, long fieldOffset, char[]? buffer, int bufferOffset, int length) =>
        _reader.GetChars(i, fieldOffset, buffer, bufferOffset, length);
    public IDataReader GetData(int i) => _reader.GetData(i);
    public string GetDataTypeName(int i) => _reader.GetDataTypeName(i);
    public DateTime GetDateTime(int i) => _reader.GetDateTime(i);
    public decimal GetDecimal(int i) => _reader.GetDecimal(i);
    public double GetDouble(int i) => _reader.GetDouble(i);
    public Type GetFieldType(int i) => _reader.GetFieldType(i);
    public float GetFloat(int i) => _reader.GetFloat(i);
    public Guid GetGuid(int i) => _reader.GetGuid(i);
    public short GetInt16(int i) => _reader.GetInt16(i);
    public int GetInt32(int i) => _reader.GetInt32(i);
    public long GetInt64(int i) => _reader.GetInt64(i);
    public string GetName(int i) => _reader.GetName(i);
    public int GetOrdinal(string name) => _reader.GetOrdinal(name);
    public string GetString(int i) => _reader.GetString(i);
    public object GetValue(int i) => _reader.GetValue(i);
    public int GetValues(object[] values) => _reader.GetValues(values);
    public bool IsDBNull(int i) => _reader.IsDBNull(i);
}
