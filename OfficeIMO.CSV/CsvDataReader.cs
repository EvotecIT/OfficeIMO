#nullable enable

using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
#if NET8_0_OR_GREATER
using CsvDataReaderTextRowSource = OfficeIMO.CSV.CsvParser.CsvTextDataReaderRowSource;
#else
using CsvDataReaderTextRowSource = OfficeIMO.CSV.ICsvDataReaderTextRowSource;
#endif

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
    private readonly CsvDataReaderTextRowSource? _textRowSource;
    private readonly CultureInfo _culture;
    private readonly IReadOnlyList<string>? _dateTimeFormats;
    private readonly IDisposable? _rowOwner;
    private readonly CsvLoadOptions? _stringRowOptions;
    private readonly object?[]? _staticColumnValues;
    private readonly int _sourceColumnCount;
    private readonly bool _useRawStringValues;
    private readonly bool _useDirectValueConversion;
    private readonly bool _rawRowsAreParsedStringsOnly;
    private readonly string? _stringNullValue;
    private Dictionary<string, int>? _ordinals;
    private object?[]? _currentRawRow;
    private IReadOnlyList<string>? _currentStringRow;
    private object?[]? _currentConvertedRow;
    private object?[]? _bufferedRawRow;
    private IReadOnlyList<string>? _bufferedStringRow;
    private bool _hasBufferedRow;
    private bool _checkedForRows;
    private bool _hasCurrentTextRow;
    private bool _closed;
    private int _rowIndex = -1;

    internal CsvDataReader(
        CsvDataColumnProjection[] columns,
        IEnumerable<object?[]> rows,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        bool rawRowsAreParsedStringsOnly = false,
        IDisposable? rowOwner = null)
    {
        _columns = columns;
        _rows = rows.GetEnumerator();
        _culture = culture;
        _dateTimeFormats = dateTimeFormats;
        _rowOwner = rowOwner;
        _useRawStringValues = CanUseRawStringValues(columns);
        _useDirectValueConversion = CanUseDirectValueConversion(columns);
        _rawRowsAreParsedStringsOnly = rawRowsAreParsedStringsOnly;
    }

    internal CsvDataReader(
        CsvDataColumnProjection[] columns,
        IEnumerable<IReadOnlyList<string>> rows,
        int sourceColumnCount,
        CsvLoadOptions options,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        IDisposable? rowOwner = null)
    {
        _columns = columns;
        _stringRows = rows.GetEnumerator();
        _sourceColumnCount = sourceColumnCount;
        _stringRowOptions = options;
        _stringNullValue = options.NullValue;
        _staticColumnValues = CaptureStaticColumnValues(_stringRowOptions.StaticColumns);
        _culture = culture;
        _dateTimeFormats = dateTimeFormats;
        _rowOwner = rowOwner;
        _useRawStringValues = CanUseRawStringValues(columns);
        _useDirectValueConversion = CanUseDirectValueConversion(columns);
    }

    internal CsvDataReader(
        CsvDataColumnProjection[] columns,
        CsvDataReaderTextRowSource rows,
        int sourceColumnCount,
        CsvLoadOptions options,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats)
    {
        _columns = columns;
        _textRowSource = rows;
        _sourceColumnCount = sourceColumnCount;
        _stringRowOptions = options;
        _stringNullValue = options.NullValue;
        _culture = culture;
        _dateTimeFormats = dateTimeFormats;
        _useRawStringValues = CanUseRawStringValues(columns);
        _useDirectValueConversion = CanUseDirectValueConversion(columns);
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
    public override bool GetBoolean(int ordinal)
    {
        EnsureOpenRow();
        if (_columns[ordinal].ConversionKind == CsvDataConversionKind.Boolean)
        {
            var rawValue = GetRawValue(ordinal);
            if (rawValue is bool boolean)
            {
                return boolean;
            }

            if (rawValue is string text)
            {
                if (bool.TryParse(text, out boolean))
                {
                    return boolean;
                }

                if (text == "0" || text == "1")
                {
                    return text == "1";
                }
            }
        }

        return (bool)GetValue(ordinal);
    }

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
    public override DateTime GetDateTime(int ordinal)
    {
        EnsureOpenRow();
        if (_columns[ordinal].ConversionKind == CsvDataConversionKind.DateTime)
        {
            var rawValue = GetRawValue(ordinal);
            if (rawValue is DateTime dateTime)
            {
                return dateTime;
            }

            if (rawValue is string text && TryParseDateTime(text, out dateTime))
            {
                return dateTime;
            }
        }

        return (DateTime)GetValue(ordinal);
    }

    /// <inheritdoc />
    public override decimal GetDecimal(int ordinal)
    {
        EnsureOpenRow();
        if (_columns[ordinal].ConversionKind == CsvDataConversionKind.Decimal)
        {
            var rawValue = GetRawValue(ordinal);
            if (rawValue is decimal decimalValue)
            {
                return decimalValue;
            }

            if (rawValue is string text)
            {
                if (ReferenceEquals(_culture, CultureInfo.InvariantCulture) &&
                    CsvDataProjectionConverter.TryParseInvariantDecimal(text, out decimalValue))
                {
                    return decimalValue;
                }

                if (decimal.TryParse(text, NumberStyles.Any, _culture, out decimalValue))
                {
                    return decimalValue;
                }
            }
        }

        return (decimal)GetValue(ordinal);
    }

    /// <inheritdoc />
    public override double GetDouble(int ordinal)
    {
        EnsureOpenRow();
        if (_columns[ordinal].ConversionKind == CsvDataConversionKind.Double)
        {
            var rawValue = GetRawValue(ordinal);
            if (rawValue is double doubleValue)
            {
                return doubleValue;
            }

            if (rawValue is string text && double.TryParse(text, NumberStyles.Any, _culture, out doubleValue))
            {
                return doubleValue;
            }
        }

        return (double)GetValue(ordinal);
    }

    /// <inheritdoc />
    public override IEnumerator GetEnumerator()
    {
        while (Read())
        {
            yield return this;
        }
    }

    /// <inheritdoc />
    [UnconditionalSuppressMessage("Trimming", "IL2073", Justification = "CSV column types are scalar conversion tokens returned through DbDataReader; OfficeIMO never activates or reflects over their public members.")]
    [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
    public override Type GetFieldType(int ordinal) => _columns[ordinal].DataType;

    /// <inheritdoc />
    public override float GetFloat(int ordinal) => (float)GetValue(ordinal);

    /// <inheritdoc />
    public override Guid GetGuid(int ordinal) => (Guid)GetValue(ordinal);

    /// <inheritdoc />
    public override short GetInt16(int ordinal)
    {
        EnsureOpenRow();
        if (_columns[ordinal].ConversionKind == CsvDataConversionKind.Int16)
        {
            var rawValue = GetRawValue(ordinal);
            if (rawValue is short int16)
            {
                return int16;
            }

            if (rawValue is string text && short.TryParse(text, NumberStyles.Any, _culture, out int16))
            {
                return int16;
            }
        }

        return (short)GetValue(ordinal);
    }

    /// <inheritdoc />
    public override int GetInt32(int ordinal)
    {
        EnsureOpenRow();
        if (_columns[ordinal].ConversionKind == CsvDataConversionKind.Int32)
        {
            var rawValue = GetRawValue(ordinal);
            if (rawValue is int int32)
            {
                return int32;
            }

            if (rawValue is string text)
            {
                if (ReferenceEquals(_culture, CultureInfo.InvariantCulture) &&
                    CsvDataProjectionConverter.TryParseInvariantInt32(text, out int32))
                {
                    return int32;
                }

                if (int.TryParse(text, NumberStyles.Any, _culture, out int32))
                {
                    return int32;
                }
            }
        }

        return (int)GetValue(ordinal);
    }

    /// <inheritdoc />
    public override long GetInt64(int ordinal)
    {
        EnsureOpenRow();
        if (_columns[ordinal].ConversionKind == CsvDataConversionKind.Int64)
        {
            var rawValue = GetRawValue(ordinal);
            if (rawValue is long int64)
            {
                return int64;
            }

            if (rawValue is string text && long.TryParse(text, NumberStyles.Any, _culture, out int64))
            {
                return int64;
            }
        }

        return (long)GetValue(ordinal);
    }

    /// <inheritdoc />
    public override string GetName(int ordinal) => _columns[ordinal].Name;

    /// <inheritdoc />
    public override int GetOrdinal(string name)
    {
        _ordinals ??= CreateOrdinalMap(_columns);
        if (_ordinals.TryGetValue(name, out var ordinal))
        {
            return ordinal;
        }

        throw new IndexOutOfRangeException(name);
    }

    /// <inheritdoc />
    public override string GetString(int ordinal)
    {
        EnsureOpenRow();
        if (_columns[ordinal].ConversionKind == CsvDataConversionKind.String)
        {
            if (_textRowSource is not null &&
                (_stringNullValue is null || !_textRowSource.IsNull(ordinal, _stringNullValue)))
            {
                var textValue = _textRowSource.GetString(ordinal);
                if (textValue.Length == 0 && _columns[ordinal].SchemaColumn?.IsRequired == true)
                {
                    CsvDataProjectionConverter.GetDirectMissingValue(_columns[ordinal], _rowIndex);
                }

                return textValue;
            }

            var rawValue = GetRawValue(ordinal);
            if (rawValue is string text)
            {
                if (text.Length == 0 && _columns[ordinal].SchemaColumn?.IsRequired == true)
                {
                    CsvDataProjectionConverter.GetDirectMissingValue(_columns[ordinal], _rowIndex);
                }

                return text;
            }
        }

        return (string)GetValue(ordinal);
    }

    /// <inheritdoc />
    public override object GetValue(int ordinal)
    {
        EnsureOpenRow();
        if (_useRawStringValues)
        {
            if (_textRowSource is not null)
            {
                return GetTextSourceRawStringValue(ordinal);
            }

            if (_rawRowsAreParsedStringsOnly && _currentRawRow is { } rawRow)
            {
                return GetParsedStringRawValue(rawRow, ordinal);
            }

            return GetRawStringValue(ordinal);
        }

        _currentConvertedRow ??= new object?[_columns.Length];
        var value = _currentConvertedRow[ordinal];
        if (value is not null)
        {
            return value;
        }

#if NET8_0_OR_GREATER
        if (_textRowSource is not null && _useDirectValueConversion)
        {
            value = ConvertDirectTextSourceValue(ordinal);
        }
        else
#endif
        {
            var rawValue = GetRawValue(ordinal);
            value = CsvDataProjectionConverter.ConvertValue(rawValue, _columns[ordinal], _rowIndex, _culture, _dateTimeFormats);
        }

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
            if (_textRowSource is not null)
            {
                return _textRowSource.CopyStringValues(values, count, _stringNullValue);
            }

            if (_currentStringRow is not null)
            {
                return GetStringRowValues(values, count);
            }

            var row = _currentRawRow!;
            if (_rawRowsAreParsedStringsOnly)
            {
                return GetParsedStringRawRowValues(row, values, count);
            }

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

        if (_useDirectValueConversion)
        {
            if (_textRowSource is not null)
            {
                return GetDirectTextSourceValues(values, count);
            }

            if (_currentStringRow is not null)
            {
                return GetDirectStringRowValues(values, count);
            }

            var row = _currentRawRow!;
            var rowValueCount = Math.Min(count, row.Length);
            for (var i = 0; i < rowValueCount; i++)
            {
                values[i] = CsvDataProjectionConverter.ConvertValue(row[i], _columns[i], _rowIndex, _culture, _dateTimeFormats);
            }

            for (var i = rowValueCount; i < count; i++)
            {
                values[i] = CsvDataProjectionConverter.ConvertValue(null, _columns[i], _rowIndex, _culture, _dateTimeFormats);
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
    public override bool IsDBNull(int ordinal)
    {
        EnsureOpenRow();
        var column = _columns[ordinal];
        if (column.ConversionKind != CsvDataConversionKind.General)
        {
            var rawValue = GetRawValue(ordinal);
            var isMissing = rawValue is null ||
                rawValue == DBNull.Value ||
                (rawValue is string { Length: 0 } &&
                    (column.ConversionKind != CsvDataConversionKind.String ||
                     column.SchemaColumn?.IsRequired == true ||
                     column.SchemaColumn?.DefaultValue is not null));
            if (!isMissing)
            {
                return false;
            }

            if (column.SchemaColumn?.IsRequired == true || column.SchemaColumn?.DefaultValue is not null)
            {
                return GetValue(ordinal) == DBNull.Value;
            }

            return true;
        }

        return GetValue(ordinal) == DBNull.Value;
    }

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
            if (_textRowSource is not null)
            {
                _hasCurrentTextRow = true;
            }
            else
            {
                _currentRawRow = _bufferedRawRow;
                _currentStringRow = _bufferedStringRow;
            }

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
        _textRowSource?.Dispose();
        _rowOwner?.Dispose();
    }

    /// <inheritdoc />
    [UnconditionalSuppressMessage("Trimming", "IL2111", Justification = "The schema table stores Type values as data and does not reflect over Type.TypeInitializer or other Type members.")]
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
            return _currentRawRow is not null || _currentStringRow is not null || _hasCurrentTextRow;
        }

        _checkedForRows = true;
        if (_textRowSource is not null)
        {
            if (!_textRowSource.Read())
            {
                return false;
            }

            _hasBufferedRow = true;
            return true;
        }

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
        if (_textRowSource is not null)
        {
            _hasCurrentTextRow = _textRowSource.Read();
            return _hasCurrentTextRow;
        }

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
        _hasCurrentTextRow = false;
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

    private static bool CanUseDirectValueConversion(CsvDataColumnProjection[] columns)
    {
        for (var i = 0; i < columns.Length; i++)
        {
            if (columns[i].ConversionKind == CsvDataConversionKind.General)
            {
                return false;
            }
        }

        return true;
    }

    private static Dictionary<string, int> CreateOrdinalMap(CsvDataColumnProjection[] columns)
    {
        var ordinals = new Dictionary<string, int>(columns.Length, StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < columns.Length; i++)
        {
            if (!ordinals.ContainsKey(columns[i].Name))
            {
                ordinals.Add(columns[i].Name, i);
            }
        }

        return ordinals;
    }

    private object GetRawStringValue(int ordinal)
    {
        if ((uint)ordinal >= (uint)_columns.Length)
        {
            throw new IndexOutOfRangeException();
        }

        if (_currentStringRow is not null)
        {
            if (ordinal < _sourceColumnCount)
            {
                if (ordinal >= _currentStringRow.Count)
                {
                    return string.Empty;
                }

                var value = _currentStringRow[ordinal];
                return _stringNullValue is not null && string.Equals(value, _stringNullValue, StringComparison.Ordinal)
                    ? DBNull.Value
                    : value;
            }

            var staticIndex = ordinal - _sourceColumnCount;
            var staticValue = _staticColumnValues is not null && staticIndex < _staticColumnValues.Length
                ? _staticColumnValues[staticIndex]
                : null;

            return staticValue switch
            {
                null => DBNull.Value,
                DBNull => DBNull.Value,
                string staticText => staticText,
                _ => CsvDataProjectionConverter.ConvertValue(staticValue, _columns[ordinal], _rowIndex, _culture, _dateTimeFormats)
            };
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
        if (_textRowSource is not null)
        {
            return _stringNullValue is not null && _textRowSource.IsNull(ordinal, _stringNullValue)
                ? null
                : _textRowSource.GetString(ordinal);
        }

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
            return _stringNullValue is not null &&
                string.Equals(value, _stringNullValue, StringComparison.Ordinal)
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
        if (_stringNullValue is null)
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
                values[i] = string.Equals(value, _stringNullValue, StringComparison.Ordinal)
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

    private int GetDirectStringRowValues(object[] values, int count)
    {
        var row = _currentStringRow!;
        var rowValueCount = Math.Min(Math.Min(count, row.Count), _sourceColumnCount);
        for (var i = 0; i < rowValueCount; i++)
        {
            var value = row[i];
            values[i] = _stringNullValue is not null && string.Equals(value, _stringNullValue, StringComparison.Ordinal)
                ? ConvertDirectStringValue(i, null)
                : ConvertDirectStringValue(i, value);
        }

        var missingCount = Math.Min(count, _sourceColumnCount);
        for (var i = rowValueCount; i < missingCount; i++)
        {
            values[i] = ConvertDirectStringValue(i, string.Empty);
        }

        for (var i = _sourceColumnCount; i < count; i++)
        {
            var staticIndex = i - _sourceColumnCount;
            var staticValue = _staticColumnValues is not null && staticIndex < _staticColumnValues.Length
                ? _staticColumnValues[staticIndex]
                : null;

            values[i] = CsvDataProjectionConverter.ConvertValue(staticValue, _columns[i], _rowIndex, _culture, _dateTimeFormats);
        }

        return count;
    }

    private object GetTextSourceRawStringValue(int ordinal)
    {
        if ((uint)ordinal >= (uint)_sourceColumnCount)
        {
            throw new IndexOutOfRangeException();
        }

        var source = _textRowSource!;
        return _stringNullValue is not null && source.IsNull(ordinal, _stringNullValue)
            ? DBNull.Value
            : source.GetString(ordinal);
    }

    private int GetDirectTextSourceValues(object[] values, int count)
    {
        var source = _textRowSource!;
        var valueCount = Math.Min(count, _sourceColumnCount);
        for (var i = 0; i < valueCount; i++)
        {
#if NET8_0_OR_GREATER
            values[i] = ConvertDirectTextSourceValue(i);
#else
            values[i] = _stringNullValue is not null && source.IsNull(i, _stringNullValue)
                ? ConvertDirectStringValue(i, null)
                : ConvertDirectStringValue(i, source.GetString(i));
#endif
        }

        for (var i = valueCount; i < count; i++)
        {
            values[i] = DBNull.Value;
        }

        return count;
    }

#if NET8_0_OR_GREATER
    private object ConvertDirectTextSourceValue(int ordinal)
    {
        var source = _textRowSource!;
        if (_stringNullValue is not null && source.IsNull(ordinal, _stringNullValue))
        {
            return CsvDataProjectionConverter.GetDirectMissingValue(_columns[ordinal], _rowIndex);
        }

        var column = _columns[ordinal];
        var text = source.GetSpan(ordinal);
        if (text.Length == 0 && column.SchemaColumn?.IsRequired == true)
        {
            return CsvDataProjectionConverter.GetDirectMissingValue(column, _rowIndex);
        }

        if (column.ConversionKind == CsvDataConversionKind.String)
        {
            return source.GetString(ordinal);
        }

        return CsvDataProjectionConverter.ConvertTextSpan(
            text,
            column,
            _rowIndex,
            _culture,
            _dateTimeFormats);
    }
#endif

    private object ConvertDirectStringValue(int ordinal, string? text)
    {
        var column = _columns[ordinal];
        if (text is null ||
            (text.Length == 0 &&
                (column.ConversionKind != CsvDataConversionKind.String || column.SchemaColumn?.IsRequired == true)))
        {
            return CsvDataProjectionConverter.GetDirectMissingValue(column, _rowIndex);
        }

        switch (column.ConversionKind)
        {
            case CsvDataConversionKind.String:
                return text;
            case CsvDataConversionKind.Int32:
                if (ReferenceEquals(_culture, CultureInfo.InvariantCulture) &&
                    CsvDataProjectionConverter.TryParseInvariantInt32(text, out var fastInt32))
                {
                    return fastInt32;
                }

                if (int.TryParse(text, NumberStyles.Any, _culture, out var int32))
                {
                    return int32;
                }

                break;
            case CsvDataConversionKind.Decimal:
                if (ReferenceEquals(_culture, CultureInfo.InvariantCulture) &&
                    CsvDataProjectionConverter.TryParseInvariantDecimal(text, out var fastDecimal))
                {
                    return fastDecimal;
                }

                if (decimal.TryParse(text, NumberStyles.Any, _culture, out var decimalValue))
                {
                    return decimalValue;
                }

                break;
            case CsvDataConversionKind.DateTime:
                if (TryParseDateTime(text, out var dateTime))
                {
                    return dateTime;
                }

                break;
            case CsvDataConversionKind.Boolean:
                if (bool.TryParse(text, out var boolean))
                {
                    return boolean;
                }

                if (text == "0" || text == "1")
                {
                    return text == "1";
                }

                break;
            case CsvDataConversionKind.Guid:
                if (Guid.TryParse(text, out var guid))
                {
                    return guid;
                }

                break;
        }

        return CsvDataProjectionConverter.ConvertValue(text, column, _rowIndex, _culture, _dateTimeFormats);
    }

    private static int GetParsedStringRawRowValues(object?[] row, object[] values, int count)
    {
        var rowValueCount = Math.Min(count, row.Length);
        Array.Copy(row, values, rowValueCount);
        for (var i = rowValueCount; i < count; i++)
        {
            values[i] = DBNull.Value;
        }

        return count;
    }

    private bool TryParseDateTime(string text, out DateTime dateTime)
    {
        if (_dateTimeFormats is { Count: > 0 } &&
            DateTime.TryParseExact(text, _dateTimeFormats as string[] ?? _dateTimeFormats.ToArray(), _culture, DateTimeStyles.None, out dateTime))
        {
            return true;
        }

        if (_dateTimeFormats is not { Count: > 0 } &&
            ReferenceEquals(_culture, CultureInfo.InvariantCulture) &&
            CsvDataProjectionConverter.TryParseDefaultInvariantDateTime(text, out dateTime))
        {
            return true;
        }

        return DateTime.TryParse(text, _culture, DateTimeStyles.None, out dateTime);
    }

    private object GetParsedStringRawValue(object?[] row, int ordinal)
    {
        if ((uint)ordinal >= (uint)_columns.Length)
        {
            throw new IndexOutOfRangeException();
        }

        return ordinal < row.Length
            ? row[ordinal] ?? DBNull.Value
            : DBNull.Value;
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

        if (_currentRawRow is null && _currentStringRow is null && !_hasCurrentTextRow)
        {
            throw new InvalidOperationException("The reader is not positioned on a row.");
        }
    }
}
