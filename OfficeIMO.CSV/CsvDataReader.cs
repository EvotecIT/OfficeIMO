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
    private readonly IEnumerator<object?[]>? _rows;
    private readonly IEnumerator<IReadOnlyList<string>>? _stringRows;
    private readonly CultureInfo _culture;
    private readonly IReadOnlyList<string>? _dateTimeFormats;
    private readonly Dictionary<string, int> _ordinals;
    private readonly CsvLoadOptions? _stringRowOptions;
    private readonly object?[]? _staticColumnValues;
    private readonly int _sourceColumnCount;
    private readonly bool _useRawStringValues;
    private object?[]? _currentRawRow;
    private IReadOnlyList<string>? _currentStringRow;
    private object?[]? _currentConvertedRow;
    private object?[]? _bufferedRawRow;
    private IReadOnlyList<string>? _bufferedStringRow;
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
        _useRawStringValues = CanUseRawStringValues(columns);

        for (var i = 0; i < columns.Length; i++)
        {
            _ordinals[columns[i].Name] = i;
        }
    }

    internal CsvDataReader(
        CsvDataColumnProjection[] columns,
        IEnumerable<IReadOnlyList<string>> rows,
        int sourceColumnCount,
        CsvLoadOptions options,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats)
    {
        _columns = columns;
        _stringRows = rows.GetEnumerator();
        _sourceColumnCount = sourceColumnCount;
        _stringRowOptions = options.Clone();
        _staticColumnValues = CaptureStaticColumnValues(_stringRowOptions.StaticColumns);
        _culture = culture;
        _dateTimeFormats = dateTimeFormats;
        _ordinals = new Dictionary<string, int>(columns.Length, StringComparer.OrdinalIgnoreCase);
        _useRawStringValues = CanUseRawStringValues(columns);

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

        if (dataOffset >= value.Length || length == 0)
        {
            return 0;
        }

        var offset = (int)dataOffset;
        var available = value.Length - offset;
        var count = Math.Min(length, available);
        if (count <= 0)
        {
            return 0;
        }

        value.CopyTo(offset, buffer, bufferOffset, count);
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
        if (_useRawStringValues)
        {
            return GetRawStringValue(ordinal);
        }

        _currentConvertedRow ??= new object?[_columns.Length];
        var value = _currentConvertedRow[ordinal];
        if (value is not null)
        {
            return value;
        }

        var rawValue = GetRawValue(ordinal);
        value = CsvDataProjectionConverter.ConvertValue(rawValue, _columns[ordinal], _rowIndex, _culture, _dateTimeFormats);
        _currentConvertedRow[ordinal] = value;
        return value;
    }

    /// <inheritdoc />
    public override int GetValues(object[] values)
    {
        EnsureOpenRow();
        var count = Math.Min(values.Length, _columns.Length);
        if (_useRawStringValues)
        {
            if (_currentStringRow is not null)
            {
                return GetStringRowValues(values, count);
            }

            var row = _currentRawRow!;
            var rowValueCount = Math.Min(count, row.Length);
            for (var i = 0; i < rowValueCount; i++)
            {
                var rawValue = row[i];
                values[i] = rawValue switch
                {
                    null => DBNull.Value,
                    DBNull => DBNull.Value,
                    string text => text,
                    _ => CsvDataProjectionConverter.ConvertValue(rawValue, _columns[i], _rowIndex, _culture, _dateTimeFormats)
                };
            }

            for (var i = rowValueCount; i < count; i++)
            {
                values[i] = DBNull.Value;
            }

            return count;
        }

        _currentConvertedRow ??= new object?[_columns.Length];
        for (var i = 0; i < count; i++)
        {
            var value = _currentConvertedRow[i];
            if (value is null)
            {
                var rawValue = GetRawValue(i);
                value = CsvDataProjectionConverter.ConvertValue(rawValue, _columns[i], _rowIndex, _culture, _dateTimeFormats);
                _currentConvertedRow[i] = value;
            }

            values[i] = value;
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
            _currentStringRow = _bufferedStringRow;
            _bufferedRawRow = null;
            _bufferedStringRow = null;
            _hasBufferedRow = false;
        }
        else
        {
            if (!MoveNextRow())
            {
                ClearCurrentRow();
                return false;
            }
        }

        ClearConvertedRow();
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
        _rows?.Dispose();
        _stringRows?.Dispose();
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

        if (_checkedForRows || _currentRawRow is not null || _currentStringRow is not null)
        {
            return _currentRawRow is not null || _currentStringRow is not null;
        }

        _checkedForRows = true;
        if (_rows is not null)
        {
            if (!_rows.MoveNext())
            {
                return false;
            }

            _bufferedRawRow = _rows.Current;
            _bufferedStringRow = null;
            _hasBufferedRow = true;
            return true;
        }

        if (_stringRows is null || !_stringRows.MoveNext())
        {
            return false;
        }

        _bufferedRawRow = null;
        _bufferedStringRow = _stringRows.Current;
        ValidateStringRowColumnCount(_bufferedStringRow);
        _hasBufferedRow = true;
        return true;
    }

    private bool MoveNextRow()
    {
        if (_rows is not null)
        {
            if (!_rows.MoveNext())
            {
                return false;
            }

            _currentRawRow = _rows.Current;
            _currentStringRow = null;
            return true;
        }

        if (_stringRows is null || !_stringRows.MoveNext())
        {
            return false;
        }

        _currentRawRow = null;
        _currentStringRow = _stringRows.Current;
        ValidateStringRowColumnCount(_currentStringRow);
        return true;
    }

    private void ClearCurrentRow()
    {
        _currentRawRow = null;
        _currentStringRow = null;
        ClearConvertedRow();
    }

    private void ClearConvertedRow()
    {
        if (_currentConvertedRow is not null)
        {
            Array.Clear(_currentConvertedRow, 0, _currentConvertedRow.Length);
        }
    }

    private static bool CanUseRawStringValues(CsvDataColumnProjection[] columns)
    {
        for (var i = 0; i < columns.Length; i++)
        {
            if (columns[i].DataType != typeof(string) || columns[i].SchemaColumn is not null)
            {
                return false;
            }
        }

        return true;
    }

    private object GetRawStringValue(int ordinal)
    {
        if ((uint)ordinal >= (uint)_columns.Length)
        {
            throw new IndexOutOfRangeException();
        }

        var rawValue = GetRawValue(ordinal);
        if (rawValue is null || rawValue == DBNull.Value)
        {
            return DBNull.Value;
        }

        if (rawValue is string text)
        {
            return text;
        }

        return CsvDataProjectionConverter.ConvertValue(rawValue, _columns[ordinal], _rowIndex, _culture, _dateTimeFormats);
    }

    private object? GetRawValue(int ordinal)
    {
        if (_currentStringRow is null)
        {
            return ordinal < _currentRawRow!.Length ? _currentRawRow[ordinal] : null;
        }

        if (ordinal < _sourceColumnCount)
        {
            if (ordinal >= _currentStringRow.Count)
            {
                return string.Empty;
            }

            var value = _currentStringRow[ordinal];
            return _stringRowOptions!.NullValue is not null &&
                string.Equals(value, _stringRowOptions.NullValue, StringComparison.Ordinal)
                    ? null
                    : value;
        }

        var staticIndex = ordinal - _sourceColumnCount;
        return _staticColumnValues is not null && staticIndex < _staticColumnValues.Length
            ? _staticColumnValues[staticIndex]
            : null;
    }

    private int GetStringRowValues(object[] values, int count)
    {
        var row = _currentStringRow!;
        var rowValueCount = Math.Min(Math.Min(count, row.Count), _sourceColumnCount);
        if (_stringRowOptions!.NullValue is null)
        {
            for (var i = 0; i < rowValueCount; i++)
            {
                values[i] = row[i];
            }
        }
        else
        {
            for (var i = 0; i < rowValueCount; i++)
            {
                var value = row[i];
                values[i] = string.Equals(value, _stringRowOptions.NullValue, StringComparison.Ordinal)
                    ? DBNull.Value
                    : value;
            }
        }

        var missingCount = Math.Min(count, _sourceColumnCount);
        for (var i = rowValueCount; i < missingCount; i++)
        {
            values[i] = string.Empty;
        }

        for (var i = _sourceColumnCount; i < count; i++)
        {
            values[i] = GetRawStringValue(i);
        }

        return count;
    }

    private void ValidateStringRowColumnCount(IReadOnlyList<string> row)
    {
        if (_stringRowOptions!.ColumnCountMismatchPolicy == CsvColumnCountMismatchPolicy.Strict &&
            row.Count != _sourceColumnCount)
        {
            throw new CsvException($"Row contains {row.Count} values but header defines {_sourceColumnCount} columns.");
        }
    }

    private static object?[]? CaptureStaticColumnValues(IReadOnlyDictionary<string, object?>? staticColumns)
    {
        if (staticColumns is null || staticColumns.Count == 0)
        {
            return null;
        }

        var values = new object?[staticColumns.Count];
        var index = 0;
        foreach (var staticColumn in staticColumns)
        {
            values[index++] = staticColumn.Value;
        }

        return values;
    }

    private void EnsureOpenRow()
    {
        if (_closed)
        {
            throw new InvalidOperationException("The reader is closed.");
        }

        if (_currentRawRow is null && _currentStringRow is null)
        {
            throw new InvalidOperationException("The reader is not positioned on a row.");
        }
    }
}
