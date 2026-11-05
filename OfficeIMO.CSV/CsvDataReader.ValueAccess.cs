using System.Runtime.CompilerServices;

namespace OfficeIMO.CSV;

public partial class CsvDataReader
{
    /// <inheritdoc />
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override string GetString(int ordinal)
    {
        if (_useDirectTextSourceStrings)
        {
            return GetDirectTextSourceString(ordinal);
        }

        return GetStringCore(ordinal);
    }

    /// <summary>Reads configured column conversions and missing-value rules.</summary>
    [MethodImpl(MethodImplOptions.NoInlining)]
    private string GetStringCore(int ordinal)
    {
        EnsureOpenRow();
        if (_columns[ordinal].ConversionKind == CsvDataConversionKind.String)
        {
            if (_textRowSource is not null &&
#if NET8_0_OR_GREATER
                _textRowSource is not ICsvAsyncDataReaderRowSource { HasStaticValues: true } &&
#endif
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
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override object GetValue(int ordinal)
    {
        if (_useDirectTextSourceStrings)
        {
            return GetDirectTextSourceString(ordinal);
        }

        return GetValueCore(ordinal);
    }

    /// <summary>Reads configured values while retaining the per-row conversion cache.</summary>
    [MethodImpl(MethodImplOptions.NoInlining)]
    private object GetValueCore(int ordinal)
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
            value = CsvDataProjectionConverter.ConvertValue(rawValue, _columns[ordinal], _rowIndex, _culture, _dateTimeFormats, _mappingErrorValuePolicy);
        }

        _currentConvertedRow[ordinal] = value;
        return value;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private string GetDirectTextSourceString(int ordinal)
    {
        if (!_hasCurrentTextRow)
        {
            ThrowInvalidTextReaderState();
        }

#if NET8_0_OR_GREATER
        if (_utf8StreamTextRowSource is not null)
        {
            return _utf8StreamTextRowSource.GetString(ordinal);
        }

        if (_streamTextRowSource is not null)
        {
            return _streamTextRowSource.GetString(ordinal);
        }
#endif
        return _textRowSource!.GetString(ordinal);
    }
}
