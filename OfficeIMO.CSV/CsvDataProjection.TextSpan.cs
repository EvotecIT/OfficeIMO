#nullable enable

using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.CSV;

internal static partial class CsvDataProjectionConverter
{
#if NET8_0_OR_GREATER
    internal static object ConvertTextSpan(
        ReadOnlySpan<char> text,
        CsvDataColumnProjection column,
        int rowIndex,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats)
    {
        if (text.Length == 0 &&
            (column.ConversionKind != CsvDataConversionKind.String || column.SchemaColumn?.IsRequired == true))
        {
            return GetDirectMissingValue(column, rowIndex);
        }

        switch (column.ConversionKind)
        {
            case CsvDataConversionKind.String:
                return text.ToString();
            case CsvDataConversionKind.Int32:
                if (TryParseInt32(text, culture, out var int32))
                {
                    return int32;
                }

                break;
            case CsvDataConversionKind.Int64 when long.TryParse(text, NumberStyles.Any, culture, out var int64):
                return int64;
            case CsvDataConversionKind.Int16 when short.TryParse(text, NumberStyles.Any, culture, out var int16):
                return int16;
            case CsvDataConversionKind.Byte when byte.TryParse(text, NumberStyles.Any, culture, out var byteValue):
                return byteValue;
            case CsvDataConversionKind.Boolean:
                if (bool.TryParse(text, out var boolean))
                {
                    return boolean;
                }

                if (text.Length == 1 && (text[0] == '0' || text[0] == '1'))
                {
                    return text[0] == '1';
                }

                break;
            case CsvDataConversionKind.Double when double.TryParse(text, NumberStyles.Any, culture, out var doubleValue):
                return doubleValue;
            case CsvDataConversionKind.Decimal when decimal.TryParse(text, NumberStyles.Any, culture, out var decimalValue):
                return decimalValue;
            case CsvDataConversionKind.Single when float.TryParse(text, NumberStyles.Any, culture, out var singleValue):
                return singleValue;
            case CsvDataConversionKind.DateTime when TryParseDateTime(text, culture, dateTimeFormats, out var dateTime):
                return dateTime;
            case CsvDataConversionKind.Guid when Guid.TryParse(text, out var guid):
                return guid;
            default:
                return ConvertValue(text.ToString(), column, rowIndex, culture, dateTimeFormats);
        }

        var value = text.ToString();
        throw new CsvException($"Column '{column.Name}' value on row {rowIndex + 1} cannot be converted to {column.DataType.Name}: Cannot parse '{value}' as {column.DataType.Name}.");
    }

    private static bool TryParseInt32(ReadOnlySpan<char> text, CultureInfo culture, out int value)
    {
        if (!ReferenceEquals(culture, CultureInfo.InvariantCulture))
        {
            return int.TryParse(text, NumberStyles.Any, culture, out value);
        }

        if (TryParseInvariantInt32(text, out value))
        {
            return true;
        }

        return int.TryParse(text, NumberStyles.Any, culture, out value);
    }

    private static bool TryParseInvariantInt32(ReadOnlySpan<char> text, out int value)
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
        }
        else if (text[0] == '+')
        {
            index = 1;
        }

        if (index == text.Length)
        {
            return false;
        }

        var result = 0u;
        var limit = negative ? 2147483648u : int.MaxValue;
        for (; index < text.Length; index++)
        {
            var digit = (uint)(text[index] - '0');
            if (digit > 9 || result > (limit - digit) / 10)
            {
                return false;
            }

            result = (result * 10) + digit;
        }

        value = negative
            ? result == 2147483648u ? int.MinValue : -(int)result
            : (int)result;
        return true;
    }

    private static bool TryParseDateTime(
        ReadOnlySpan<char> text,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        out DateTime dateTime)
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

    private static bool TryParseDefaultInvariantDateTime(ReadOnlySpan<char> text, out DateTime dateTime)
    {
        dateTime = default;
        if (text.Length != DefaultInvariantDateTimeFormat.Length ||
            text[2] != '/' ||
            text[5] != '/' ||
            text[10] != ' ' ||
            text[13] != ':' ||
            text[16] != ':' ||
            !TryParseTwoDigits(text, 0, out var month) ||
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

    private static bool TryParseTwoDigits(ReadOnlySpan<char> text, int offset, out int value)
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

    private static bool TryParseFourDigits(ReadOnlySpan<char> text, int offset, out int value)
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
#endif
}
