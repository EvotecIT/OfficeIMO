#nullable enable

using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.CSV;

internal static partial class CsvDataProjectionConverter
{
#if NET8_0_OR_GREATER
    private static readonly decimal[] InvariantDecimalPowers =
    {
        1m,
        10m,
        100m,
        1_000m,
        10_000m,
        100_000m,
        1_000_000m,
        10_000_000m,
        100_000_000m,
        1_000_000_000m,
        10_000_000_000m,
        100_000_000_000m,
        1_000_000_000_000m,
        10_000_000_000_000m,
        100_000_000_000_000m,
        1_000_000_000_000_000m,
        10_000_000_000_000_000m,
        100_000_000_000_000_000m,
        1_000_000_000_000_000_000m,
        10_000_000_000_000_000_000m
    };

    internal static object ConvertTextSpan(
        ReadOnlySpan<char> text,
        CsvDataColumnProjection column,
        int rowIndex,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        DataMappingErrorValuePolicy errorValuePolicy = DataMappingErrorValuePolicy.Include)
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
                return ConvertValue(text.ToString(), column, rowIndex, culture, dateTimeFormats, errorValuePolicy);
        }

        var value = text.ToString();
        throw CreateConversionException(column, rowIndex, value, errorValuePolicy);
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

    internal static bool TryParseInvariantInt32(ReadOnlySpan<char> text, out int value)
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

    internal static bool TryParseDateTime(
        ReadOnlySpan<char> text,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        out DateTime dateTime)
    {
        if (dateTimeFormats is { Count: > 0 } &&
            DateTime.TryParseExact(text, dateTimeFormats as string[] ?? dateTimeFormats.ToArray(), culture, DateTimeStyles.RoundtripKind, out dateTime))
        {
            return true;
        }

        if (dateTimeFormats is not { Count: > 0 } &&
            ReferenceEquals(culture, CultureInfo.InvariantCulture) &&
            TryParseDefaultInvariantDateTime(text, out dateTime))
        {
            return true;
        }

        return DateTime.TryParse(text, culture, DateTimeStyles.RoundtripKind, out dateTime);
    }

    internal static bool TryParseInvariantDecimal(ReadOnlySpan<char> text, out decimal value)
    {
        value = 0m;
        if (text.Length == 0)
        {
            return false;
        }

        int index = 0;
        bool negative = false;
        if (text[0] == '-' || text[0] == '+')
        {
            negative = text[0] == '-';
            index = 1;
            if (index == text.Length)
            {
                return false;
            }
        }

        bool sawDigit = false;
        bool sawDecimalPoint = false;
        int fractionalDigits = 0;
        ulong significand = 0;
        for (; index < text.Length; index++)
        {
            char current = text[index];
            if (current == '.')
            {
                if (sawDecimalPoint)
                {
                    return false;
                }

                sawDecimalPoint = true;
                continue;
            }

            uint digit = (uint)(current - '0');
            if (digit > 9)
            {
                return false;
            }

            if (significand > (ulong.MaxValue - digit) / 10UL)
            {
                return false;
            }

            sawDigit = true;
            significand = (significand * 10UL) + digit;
            if (sawDecimalPoint)
            {
                fractionalDigits++;
                if (fractionalDigits >= InvariantDecimalPowers.Length)
                {
                    return false;
                }
            }
        }

        if (!sawDigit)
        {
            value = 0m;
            return false;
        }

        value = new decimal(
            unchecked((int)(uint)significand),
            unchecked((int)(uint)(significand >> 32)),
            0,
            negative,
            (byte)fractionalDigits);
        return true;
    }

    private static bool TryParseDefaultInvariantDateTime(ReadOnlySpan<char> text, out DateTime dateTime)
    {
        dateTime = default;
        // This is the round-trip format emitted by CsvSaveOptions for UTC and
        // unspecified DateTime values. Offset-bearing local values deliberately
        // retain the framework fallback so RoundtripKind semantics stay authoritative.
        if ((text.Length == 27 || text.Length == 28) &&
            text[4] == '-' &&
            text[7] == '-' &&
            text[10] == 'T' &&
            text[13] == ':' &&
            text[16] == ':' &&
            text[19] == '.' &&
            (text.Length == 27 || text[27] == 'Z') &&
            TryParseFourDigits(text, 0, out var roundTripYear) &&
            TryParseTwoDigits(text, 5, out var roundTripMonth) &&
            TryParseTwoDigits(text, 8, out var roundTripDay) &&
            TryParseTwoDigits(text, 11, out var roundTripHour) &&
            TryParseTwoDigits(text, 14, out var roundTripMinute) &&
            TryParseTwoDigits(text, 17, out var roundTripSecond) &&
            TryParseSevenDigits(text, 20, out var fractionalTicks))
        {
            try
            {
                var kind = text.Length == 28 ? DateTimeKind.Utc : DateTimeKind.Unspecified;
                dateTime = new DateTime(
                    roundTripYear,
                    roundTripMonth,
                    roundTripDay,
                    roundTripHour,
                    roundTripMinute,
                    roundTripSecond,
                    kind).AddTicks(fractionalTicks);
                return true;
            }
            catch (ArgumentOutOfRangeException)
            {
                dateTime = default;
                return false;
            }
        }

        if (text.Length == 10 &&
            text[4] == '-' &&
            text[7] == '-' &&
            TryParseFourDigits(text, 0, out var isoYear) &&
            TryParseTwoDigits(text, 5, out var isoMonth) &&
            TryParseTwoDigits(text, 8, out var isoDay))
        {
            try
            {
                dateTime = new DateTime(isoYear, isoMonth, isoDay);
                return true;
            }
            catch (ArgumentOutOfRangeException)
            {
                dateTime = default;
                return false;
            }
        }

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

    private static bool TryParseSevenDigits(ReadOnlySpan<char> text, int offset, out long value)
    {
        value = 0;
        for (var index = 0; index < 7; index++)
        {
            var digit = text[offset + index] - '0';
            if ((uint)digit > 9)
            {
                value = 0;
                return false;
            }

            value = (value * 10) + digit;
        }

        return true;
    }
#endif
}
