using System;
using System.Data;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services;

/// <summary>Preserves streaming IDataReader behavior while normalizing complex Office table cell values.</summary>
internal sealed class PowerShellNormalizingDataReader : IDataReader
{
    private readonly IDataReader _reader;
    private readonly PowerShellObjectNormalizerOptions _options;

    internal PowerShellNormalizingDataReader(IDataReader reader, PowerShellObjectNormalizerOptions options)
    {
        _reader = reader ?? throw new ArgumentNullException(nameof(reader));
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    public object this[int i] => GetValue(i);
    public object this[string name] => GetValue(GetOrdinal(name));
    public int Depth => _reader.Depth;
    public bool IsClosed => _reader.IsClosed;
    public int RecordsAffected => _reader.RecordsAffected;
    public int FieldCount => _reader.FieldCount;

    public void Close() => _reader.Close();
    public void Dispose() => _reader.Dispose();
    public bool GetBoolean(int i) => _reader.GetBoolean(i);
    public byte GetByte(int i) => _reader.GetByte(i);
    public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) => _reader.GetBytes(i, fieldOffset, buffer, bufferoffset, length);
    public char GetChar(int i) => _reader.GetChar(i);
    public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length) => _reader.GetChars(i, fieldoffset, buffer, bufferoffset, length);
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
    public DataTable? GetSchemaTable() => _reader.GetSchemaTable();
    public string GetString(int i) => _reader.GetString(i);

    public object GetValue(int i)
    {
        var value = _reader.GetValue(i);
        if (value == DBNull.Value)
        {
            return value;
        }

        return PowerShellObjectNormalizer.NormalizeCellValueForTable(value, _options) ?? DBNull.Value;
    }

    public int GetValues(object[] values)
    {
        if (values == null)
        {
            throw new ArgumentNullException(nameof(values));
        }

        var count = Math.Min(values.Length, FieldCount);
        for (var index = 0; index < count; index++)
        {
            values[index] = GetValue(index);
        }

        return count;
    }

    public bool IsDBNull(int i) => _reader.IsDBNull(i);
    public bool NextResult() => _reader.NextResult();
    public bool Read() => _reader.Read();
}
