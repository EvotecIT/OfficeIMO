#nullable enable

using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.CSV;

internal readonly struct CsvDataColumnProjection
{
    public CsvDataColumnProjection(string name, Type dataType, CsvSchemaColumn? schemaColumn)
    {
        Name = name;
        DataType = dataType;
        SchemaColumn = schemaColumn;
        ConversionKind = ResolveConversionKind(dataType, schemaColumn);
    }

    public string Name { get; }

    public Type DataType { get; }

    public CsvSchemaColumn? SchemaColumn { get; }

    public CsvDataConversionKind ConversionKind { get; }

    private static CsvDataConversionKind ResolveConversionKind(Type dataType, CsvSchemaColumn? schemaColumn)
    {
        if (schemaColumn is { Converter: not null } ||
            schemaColumn is { DefaultValue: not null })
        {
            return CsvDataConversionKind.General;
        }

        if (dataType == typeof(string))
        {
            return CsvDataConversionKind.String;
        }

        if (dataType == typeof(int))
        {
            return CsvDataConversionKind.Int32;
        }

        if (dataType == typeof(long))
        {
            return CsvDataConversionKind.Int64;
        }

        if (dataType == typeof(short))
        {
            return CsvDataConversionKind.Int16;
        }

        if (dataType == typeof(byte))
        {
            return CsvDataConversionKind.Byte;
        }

        if (dataType == typeof(bool))
        {
            return CsvDataConversionKind.Boolean;
        }

        if (dataType == typeof(double))
        {
            return CsvDataConversionKind.Double;
        }

        if (dataType == typeof(decimal))
        {
            return CsvDataConversionKind.Decimal;
        }

        if (dataType == typeof(float))
        {
            return CsvDataConversionKind.Single;
        }

        if (dataType == typeof(DateTime))
        {
            return CsvDataConversionKind.DateTime;
        }

        if (dataType == typeof(Guid))
        {
            return CsvDataConversionKind.Guid;
        }

        return CsvDataConversionKind.General;
    }
}

internal enum CsvDataConversionKind
{
    General,
    String,
    Int32,
    Int64,
    Int16,
    Byte,
    Boolean,
    Double,
    Decimal,
    Single,
    DateTime,
    Guid
}

internal static partial class CsvDataProjectionConverter
{
    private const string DefaultInvariantDateTimeFormat = "MM/dd/yyyy HH:mm:ss";

    public static object ConvertValue(
        object? value,
        CsvDataColumnProjection column,
        int rowIndex,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats)
    {
        if (IsMissingValue(value, column))
        {
            if (column.SchemaColumn?.DefaultValue is not null)
            {
                value = column.SchemaColumn.DefaultValue;
            }
            else if (column.SchemaColumn?.IsRequired == true)
            {
                throw new CsvException($"Column '{column.Name}' is required but row {rowIndex + 1} has no value.");
            }
            else
            {
                return DBNull.Value;
            }
        }

        if (column.ConversionKind != CsvDataConversionKind.General &&
            TryConvertFast(value, column, rowIndex, culture, dateTimeFormats, out var fastValue))
        {
            return fastValue;
        }

        if (column.SchemaColumn?.Converter is { } converter)
        {
            value = ConvertWithCustomConverter(value, converter, column.Name, rowIndex, culture);
            if (IsMissingValue(value, column))
            {
                if (column.SchemaColumn.DefaultValue is not null)
                {
                    value = column.SchemaColumn.DefaultValue;
                }
                else if (column.SchemaColumn.IsRequired)
                {
                    throw new CsvException($"Column '{column.Name}' is required but row {rowIndex + 1} has no value.");
                }
                else
                {
                    return DBNull.Value;
                }
            }
        }

        if (column.DataType.IsInstanceOfType(value))
        {
            return value!;
        }

        if (!CsvValueConverter.TryConvert(value, column.DataType, culture, dateTimeFormats, out var converted, out var error))
        {
            throw new CsvException($"Column '{column.Name}' value on row {rowIndex + 1} cannot be converted to {column.DataType.Name}: {error}");
        }

        return converted ?? DBNull.Value;
    }

    internal static object GetDirectMissingValue(CsvDataColumnProjection column, int rowIndex)
    {
        if (column.SchemaColumn?.IsRequired == true)
        {
            throw new CsvException($"Column '{column.Name}' is required but row {rowIndex + 1} has no value.");
        }

        return DBNull.Value;
    }

    private static bool TryConvertFast(
        object? value,
        CsvDataColumnProjection column,
        int rowIndex,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        out object converted)
    {
        converted = DBNull.Value;
        if (value is null || value == DBNull.Value)
        {
            return true;
        }

        if (value is not string text)
        {
            if (column.DataType.IsInstanceOfType(value))
            {
                converted = value;
                return true;
            }

            return false;
        }

        if (text.Length == 0 && column.ConversionKind != CsvDataConversionKind.String)
        {
            return true;
        }

        switch (column.ConversionKind)
        {
            case CsvDataConversionKind.String:
                converted = text;
                return true;
            case CsvDataConversionKind.Int32:
                if (ReferenceEquals(culture, CultureInfo.InvariantCulture) &&
                    TryParseInvariantInt32(text, out var fastInt32))
                {
                    converted = fastInt32;
                    return true;
                }

                if (int.TryParse(text, NumberStyles.Any, culture, out var int32))
                {
                    converted = int32;
                    return true;
                }

                break;
            case CsvDataConversionKind.Int64 when long.TryParse(text, NumberStyles.Any, culture, out var int64):
                converted = int64;
                return true;
            case CsvDataConversionKind.Int16 when short.TryParse(text, NumberStyles.Any, culture, out var int16):
                converted = int16;
                return true;
            case CsvDataConversionKind.Byte when byte.TryParse(text, NumberStyles.Any, culture, out var byteValue):
                converted = byteValue;
                return true;
            case CsvDataConversionKind.Boolean:
                if (bool.TryParse(text, out var boolean))
                {
                    converted = boolean;
                    return true;
                }

                if (text == "0" || text == "1")
                {
                    converted = text == "1";
                    return true;
                }

                break;
            case CsvDataConversionKind.Double when double.TryParse(text, NumberStyles.Any, culture, out var doubleValue):
                converted = doubleValue;
                return true;
            case CsvDataConversionKind.Decimal:
                if (ReferenceEquals(culture, CultureInfo.InvariantCulture) &&
                    TryParseInvariantDecimal(text, out var fastDecimal))
                {
                    converted = fastDecimal;
                    return true;
                }

                if (decimal.TryParse(text, NumberStyles.Any, culture, out var decimalValue))
                {
                    converted = decimalValue;
                    return true;
                }

                break;
            case CsvDataConversionKind.Single when float.TryParse(text, NumberStyles.Any, culture, out var singleValue):
                converted = singleValue;
                return true;
            case CsvDataConversionKind.DateTime:
                if (TryParseDateTime(text, culture, dateTimeFormats, out var dateTime))
                {
                    converted = dateTime;
                    return true;
                }

                break;
            case CsvDataConversionKind.Guid when Guid.TryParse(text, out var guid):
                converted = guid;
                return true;
        }

        throw new CsvException($"Column '{column.Name}' value on row {rowIndex + 1} cannot be converted to {column.DataType.Name}: Cannot parse '{text}' as {column.DataType.Name}.");
    }

    private static bool TryParseDateTime(string text, CultureInfo culture, IReadOnlyList<string>? dateTimeFormats, out DateTime dateTime)
    {
        if (dateTimeFormats is { Count: > 0 } &&
            DateTime.TryParseExact(text, dateTimeFormats as string[] ?? dateTimeFormats.ToArray(), culture, DateTimeStyles.None, out dateTime))
        {
            return true;
        }

        if (dateTimeFormats is not { Count: > 0 } &&
            ReferenceEquals(culture, CultureInfo.InvariantCulture) &&
            TryParseDefaultInvariantDateTime(text, out dateTime))
        {
            return true;
        }

        return DateTime.TryParse(text, culture, DateTimeStyles.None, out dateTime);
    }

    internal static bool TryParseDefaultInvariantDateTime(string text, out DateTime dateTime)
    {
        dateTime = default;
        if (text.Length != DefaultInvariantDateTimeFormat.Length ||
            text[2] != '/' ||
            text[5] != '/' ||
            text[10] != ' ' ||
            text[13] != ':' ||
            text[16] != ':')
        {
            return false;
        }

        if (!TryParseTwoDigits(text, 0, out var month) ||
            !TryParseTwoDigits(text, 3, out var day) ||
            !TryParseFourDigits(text, 6, out var year) ||
            !TryParseTwoDigits(text, 11, out var hour) ||
            !TryParseTwoDigits(text, 14, out var minute) ||
            !TryParseTwoDigits(text, 17, out var second))
        {
            return false;
        }

        try
        {
            dateTime = new DateTime(year, month, day, hour, minute, second);
            return true;
        }
        catch (ArgumentOutOfRangeException)
        {
            dateTime = default;
            return false;
        }
    }

    internal static bool TryParseInvariantInt32(string text, out int value)
    {
        value = 0;
        if (text.Length == 0)
        {
            return false;
        }

        var index = 0;
        var negative = false;
        if (text[0] == '-')
        {
            negative = true;
            index = 1;
            if (index == text.Length)
            {
                return false;
            }
        }
        else if (text[0] == '+')
        {
            index = 1;
            if (index == text.Length)
            {
                return false;
            }
        }

        var result = 0;
        for (; index < text.Length; index++)
        {
            var digit = text[index] - '0';
            if ((uint)digit > 9)
            {
                return false;
            }

            if (result > (int.MaxValue - digit) / 10)
            {
                return false;
            }

            result = (result * 10) + digit;
        }

        value = negative ? -result : result;
        return true;
    }

    internal static bool TryParseInvariantDecimal(string text, out decimal value)
    {
        value = 0m;
        if (text.Length == 0)
        {
            return false;
        }

        var index = 0;
        var negative = false;
        if (text[0] == '-')
        {
            negative = true;
            index = 1;
            if (index == text.Length)
            {
                return false;
            }
        }
        else if (text[0] == '+')
        {
            index = 1;
            if (index == text.Length)
            {
                return false;
            }
        }

        var sawDigit = false;
        var sawDecimalPoint = false;
        var scale = 1m;
        try
        {
            for (; index < text.Length; index++)
            {
                var current = text[index];
                if (current == '.')
                {
                    if (sawDecimalPoint)
                    {
                        return false;
                    }

                    sawDecimalPoint = true;
                    continue;
                }

                var digit = current - '0';
                if ((uint)digit > 9)
                {
                    return false;
                }

                sawDigit = true;
                if (sawDecimalPoint)
                {
                    scale *= 10m;
                }

                value = (value * 10m) + digit;
            }
        }
        catch (OverflowException)
        {
            value = 0m;
            return false;
        }

        if (!sawDigit)
        {
            value = 0m;
            return false;
        }

        if (scale != 1m)
        {
            value /= scale;
        }

        if (negative)
        {
            value = -value;
        }

        return true;
    }

    private static bool TryParseTwoDigits(string text, int offset, out int value)
    {
        var tens = text[offset] - '0';
        var ones = text[offset + 1] - '0';
        if ((uint)tens > 9 || (uint)ones > 9)
        {
            value = 0;
            return false;
        }

        value = tens * 10 + ones;
        return true;
    }

    private static bool TryParseFourDigits(string text, int offset, out int value)
    {
        value = 0;
        for (var i = 0; i < 4; i++)
        {
            var digit = text[offset + i] - '0';
            if ((uint)digit > 9)
            {
                return false;
            }

            value = (value * 10) + digit;
        }

        return true;
    }

    private static object? ConvertWithCustomConverter(
        object? value,
        Func<object?, CultureInfo, object?> converter,
        string columnName,
        int rowIndex,
        CultureInfo culture)
    {
        try
        {
            return converter(value, culture);
        }
        catch (Exception ex) when (ex is not CsvException)
        {
            throw new CsvException($"Column '{columnName}' custom converter failed on row {rowIndex + 1}: {ex.Message}", ex);
        }
    }

    private static bool IsMissingValue(object? value, CsvDataColumnProjection column) =>
        value is null ||
        value == DBNull.Value ||
        (value is string { Length: 0 } &&
            (column.DataType != typeof(string) ||
             column.SchemaColumn?.IsRequired == true ||
             column.SchemaColumn?.DefaultValue is not null));
}

internal static class CsvDataProjectionBuilder
{
    public static CsvDataColumnProjection[] Create(
        IReadOnlyList<string> header,
        IReadOnlyDictionary<string, CsvSchemaColumn>? schemaColumns)
    {
        if (schemaColumns is not null)
        {
            var headerColumns = new HashSet<string>(header, StringComparer.OrdinalIgnoreCase);
            foreach (var schemaColumn in schemaColumns.Values)
            {
                if (schemaColumn.IsRequired && !headerColumns.Contains(schemaColumn.Name))
                {
                    throw new CsvException($"Required column '{schemaColumn.Name}' is missing.");
                }
            }
        }

        var columns = new CsvDataColumnProjection[header.Count];
        for (var i = 0; i < columns.Length; i++)
        {
            var columnName = header[i];
            CsvSchemaColumn? schemaColumn = null;
            schemaColumns?.TryGetValue(columnName, out schemaColumn);
            columns[i] = new CsvDataColumnProjection(columnName, ResolveDataColumnType(schemaColumn?.DataType), schemaColumn);
        }

        return columns;
    }

    public static CsvDataColumnProjection[] CreateByOrdinal(
        IReadOnlyList<string> header,
        IReadOnlyList<CsvSchemaColumn> schemaColumns)
    {
        if (header.Count != schemaColumns.Count)
        {
            throw new ArgumentException("Positional schema columns must match the header count.", nameof(schemaColumns));
        }

        var columns = new CsvDataColumnProjection[header.Count];
        for (var i = 0; i < columns.Length; i++)
        {
            var schemaColumn = schemaColumns[i];
            if (!string.Equals(header[i], schemaColumn.Name, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Positional schema columns must match the header order.", nameof(schemaColumns));
            }

            columns[i] = new CsvDataColumnProjection(
                header[i],
                ResolveDataColumnType(schemaColumn.DataType),
                schemaColumn);
        }

        return columns;
    }

    public static Type ResolveDataColumnType(Type? dataType)
    {
        if (dataType is null)
        {
            return typeof(string);
        }

        return Nullable.GetUnderlyingType(dataType) ?? dataType;
    }
}
