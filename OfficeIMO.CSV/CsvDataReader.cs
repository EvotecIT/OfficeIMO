#nullable enable

using System.Collections;
using System.Data;
using System.Data.Common;
using System.Globalization;

namespace OfficeIMO.CSV;

/// <summary>
/// Forward-only reader for CSV rows projected through an optional schema.
/// </summary>
public sealed class CsvDataReader : DbDataReader
{
    private const string IsReadOnlyColumn = "IsReadOnly";
    private const string IsRowVersionColumn = "IsRowVersion";
    private const string IsAutoIncrementColumn = "IsAutoIncrement";
    private const string BaseCatalogNameColumn = "BaseCatalogName";
    private readonly CsvDataColumnProjection[] _columns;
    private readonly IEnumerator<object?[]> _rows;
    private readonly CultureInfo _culture;
    private readonly IReadOnlyList<string>? _dateTimeFormats;
    private readonly Dictionary<string, int> _ordinals;
    private object?[]? _currentRawRow;
    private object?[]? _currentConvertedRow;
    private object?[]? _bufferedRawRow;
    private bool _hasBufferedRow;
    private bool _checkedForRows;
    private bool _closed;
    private int _rowIndex = -1;

    internal CsvDataReader(
        CsvDataColumnProjection[] columns,
        IEnumerable<object?[]> rows,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats)
    {
        _columns = columns;
        _rows = rows.GetEnumerator();
        _culture = culture;
        _dateTimeFormats = dateTimeFormats;
        _ordinals = new Dictionary<string, int>(columns.Length, StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < columns.Length; i++)
        {
            _ordinals[columns[i].Name] = i;
        }
    }

    /// <inheritdoc />
    public override object this[int ordinal] => GetValue(ordinal);

    /// <inheritdoc />
    public override object this[string name] => GetValue(GetOrdinal(name));

    /// <inheritdoc />
    public override int Depth => 0;

    /// <inheritdoc />
    public override int FieldCount => _columns.Length;

    /// <inheritdoc />
    public override bool HasRows => !_closed && EnsureBufferedRow();

    /// <inheritdoc />
    public override bool IsClosed => _closed;

    /// <inheritdoc />
    public override int RecordsAffected => -1;

    /// <inheritdoc />
    public override bool GetBoolean(int ordinal) => (bool)GetValue(ordinal);

    /// <inheritdoc />
    public override byte GetByte(int ordinal) => (byte)GetValue(ordinal);

    /// <inheritdoc />
    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) =>
        throw new NotSupportedException("CSV fields are exposed as scalar values.");

    /// <inheritdoc />
    public override char GetChar(int ordinal) => (char)GetValue(ordinal);

    /// <inheritdoc />
    public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length)
    {
        var value = Convert.ToString(GetValue(ordinal), _culture) ?? string.Empty;
        if (buffer is null)
        {
            return value.Length;
        }

        var available = Math.Max(0, value.Length - (int)dataOffset);
        var count = Math.Min(length, available);
        value.CopyTo((int)dataOffset, buffer, bufferOffset, count);
        return count;
    }

    /// <inheritdoc />
    public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;

    /// <inheritdoc />
    public override DateTime GetDateTime(int ordinal) => (DateTime)GetValue(ordinal);

    /// <inheritdoc />
    public override decimal GetDecimal(int ordinal) => (decimal)GetValue(ordinal);

    /// <inheritdoc />
    public override double GetDouble(int ordinal) => (double)GetValue(ordinal);

    /// <inheritdoc />
    public override IEnumerator GetEnumerator()
    {
        while (Read())
        {
            yield return this;
        }
    }

    /// <inheritdoc />
    public override Type GetFieldType(int ordinal) => _columns[ordinal].DataType;

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
    public override string GetName(int ordinal) => _columns[ordinal].Name;

    /// <inheritdoc />
    public override int GetOrdinal(string name)
    {
        if (_ordinals.TryGetValue(name, out var ordinal))
        {
            return ordinal;
        }

        throw new IndexOutOfRangeException(name);
    }

    /// <inheritdoc />
    public override string GetString(int ordinal) => (string)GetValue(ordinal);

    /// <inheritdoc />
    public override object GetValue(int ordinal)
    {
        EnsureOpenRow();
        _currentConvertedRow ??= new object?[_columns.Length];
        var value = _currentConvertedRow[ordinal];
        if (value is not null)
        {
            return value;
        }

        var rawValue = ordinal < _currentRawRow!.Length ? _currentRawRow[ordinal] : null;
        value = CsvDataProjectionConverter.ConvertValue(rawValue, _columns[ordinal], _rowIndex, _culture, _dateTimeFormats);
        _currentConvertedRow[ordinal] = value;
        return value;
    }

    /// <inheritdoc />
    public override int GetValues(object[] values)
    {
        EnsureOpenRow();
        var count = Math.Min(values.Length, _columns.Length);
        for (var i = 0; i < count; i++)
        {
            values[i] = GetValue(i);
        }

        return count;
    }

    /// <inheritdoc />
    public override bool IsDBNull(int ordinal) => GetValue(ordinal) == DBNull.Value;

    /// <inheritdoc />
    public override bool NextResult() => false;

    /// <inheritdoc />
    public override bool Read()
    {
        if (_closed)
        {
            return false;
        }

        if (_hasBufferedRow)
        {
            _currentRawRow = _bufferedRawRow;
            _bufferedRawRow = null;
            _hasBufferedRow = false;
        }
        else
        {
            if (!_rows.MoveNext())
            {
                _currentRawRow = null;
                _currentConvertedRow = null;
                return false;
            }

            _currentRawRow = _rows.Current;
        }

        _currentConvertedRow = null;
        _rowIndex++;
        return true;
    }

    /// <inheritdoc />
    public override void Close()
    {
        if (_closed)
        {
            return;
        }

        _closed = true;
        _rows.Dispose();
    }

    /// <inheritdoc />
    public override DataTable GetSchemaTable()
    {
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

        for (var i = 0; i < _columns.Length; i++)
        {
            var column = _columns[i];
            var row = schema.NewRow();
            row[SchemaTableColumn.ColumnName] = column.Name;
            row[SchemaTableColumn.ColumnOrdinal] = i;
            row[SchemaTableColumn.ColumnSize] = -1;
            row[SchemaTableColumn.NumericPrecision] = DBNull.Value;
            row[SchemaTableColumn.NumericScale] = DBNull.Value;
            row[SchemaTableColumn.DataType] = column.DataType;
            row[SchemaTableColumn.ProviderType] = 0;
            row[SchemaTableColumn.IsLong] = false;
            row[SchemaTableColumn.AllowDBNull] = !(column.SchemaColumn?.IsRequired == true);
            row[IsReadOnlyColumn] = true;
            row[IsRowVersionColumn] = false;
            row[SchemaTableColumn.IsUnique] = false;
            row[SchemaTableColumn.IsKey] = false;
            row[IsAutoIncrementColumn] = false;
            row[SchemaTableColumn.BaseSchemaName] = DBNull.Value;
            row[BaseCatalogNameColumn] = DBNull.Value;
            row[SchemaTableColumn.BaseTableName] = DBNull.Value;
            row[SchemaTableColumn.BaseColumnName] = column.Name;
            schema.Rows.Add(row);
        }

        return schema;
    }

    /// <inheritdoc />
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            Close();
        }

        base.Dispose(disposing);
    }

    private bool EnsureBufferedRow()
    {
        if (_hasBufferedRow)
        {
            return true;
        }

        if (_checkedForRows || _currentRawRow is not null)
        {
            return _currentRawRow is not null;
        }

        _checkedForRows = true;
        if (!_rows.MoveNext())
        {
            return false;
        }

        _bufferedRawRow = _rows.Current;
        _hasBufferedRow = true;
        return true;
    }

    private void EnsureOpenRow()
    {
        if (_closed)
        {
            throw new InvalidOperationException("The reader is closed.");
        }

        if (_currentRawRow is null)
        {
            throw new InvalidOperationException("The reader is not positioned on a row.");
        }
    }
}
