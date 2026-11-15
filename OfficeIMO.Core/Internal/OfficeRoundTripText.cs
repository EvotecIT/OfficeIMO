using System.Globalization;
#if NETCOREAPP
using CharBuffer = System.Span<char>;
#else
using CharBuffer = char[];
#endif

namespace OfficeIMO.Core.Internal;

/// <summary>
/// Text that parses back to the identical value: ISO 8601 dates and times with every non-zero fraction digit and
/// the kind or offset, shortest round-trip floating-point numbers and base64 binary. Text formats such as CSV use
/// it as their default so that writing and reading a value never loses data.
/// </summary>
internal static class OfficeRoundTripText
{
    /// <summary>ISO 8601 date and time; the fraction is left out when zero and trimmed of trailing zeros, and <c>K</c> adds <c>Z</c> or the local offset.</summary>
    internal const string DateTimeFormat = "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK";

    /// <summary>ISO 8601 date, time and offset.</summary>
    internal const string DateTimeOffsetFormat = "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFzzz";

    /// <summary>ISO 8601 date.</summary>
    internal const string DateOnlyFormat = "yyyy-MM-dd";

    /// <summary>ISO 8601 time of day with every non-zero fraction digit.</summary>
    internal const string TimeOnlyFormat = "HH:mm:ss.FFFFFFF";

    /// <summary>The longest date and time text written here: a seven-digit fraction and an offset.</summary>
    internal const int MaximumDateTimeLength = 33;

    [ThreadStatic]
    private static char[]? t_buffer;

    /// <summary>Formats the values whose usual text loses data; false leaves any other value to the caller.</summary>
    internal static bool TryFormat(object value, CultureInfo culture, out string text)
    {
        switch (value)
        {
            case DateTime dateTime:
                text = Format(dateTime);
                return true;
            case DateTimeOffset dateTimeOffset:
                text = Format(dateTimeOffset);
                return true;
            case double number:
                text = Format(number, culture);
                return true;
            case float number:
                text = Format(number, culture);
                return true;
            case byte[] bytes:
                text = Convert.ToBase64String(bytes);
                return true;
#if NET6_0_OR_GREATER
            case DateOnly dateOnly:
                text = dateOnly.ToString(DateOnlyFormat, CultureInfo.InvariantCulture);
                return true;
            case TimeOnly timeOnly:
                text = timeOnly.ToString(TimeOnlyFormat, CultureInfo.InvariantCulture);
                return true;
#endif
            default:
                text = string.Empty;
                return false;
        }
    }

    /// <summary>Formats a date and time as <see cref="DateTimeFormat"/> text.</summary>
    internal static string Format(DateTime value)
    {
        char[] buffer = Buffer;
        return new string(buffer, 0, Write(value, buffer));
    }

    /// <summary>Formats a date, time and offset as <see cref="DateTimeOffsetFormat"/> text.</summary>
    internal static string Format(DateTimeOffset value)
    {
        char[] buffer = Buffer;
        return new string(buffer, 0, Write(value, buffer));
    }

    /// <summary>
    /// The shortest text that parses back to <paramref name="value"/>. .NET Core 3.0 and later write it by default;
    /// .NET Framework writes 15 digits by default, and its "R" format can return text that parses to a neighbouring
    /// double, so the text is checked and widened to 17 digits when needed.
    /// </summary>
    internal static string Format(double value, CultureInfo culture)
    {
#if NETCOREAPP3_0_OR_GREATER
        return value.ToString(culture);
#else
        string text = value.ToString("R", culture);
        if (double.IsNaN(value) || double.IsInfinity(value) ||
            (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, culture, out double parsed) &&
             BitConverter.DoubleToInt64Bits(parsed) == BitConverter.DoubleToInt64Bits(value)))
        {
            return text;
        }

        return value.ToString("G17", culture);
#endif
    }

    /// <inheritdoc cref="Format(double, CultureInfo)"/>
    internal static string Format(float value, CultureInfo culture)
    {
#if NETCOREAPP3_0_OR_GREATER
        return value.ToString(culture);
#else
        string text = value.ToString("R", culture);
        if (float.IsNaN(value) || float.IsInfinity(value) ||
            (float.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, culture, out float parsed) && parsed.Equals(value)))
        {
            return text;
        }

        return value.ToString("G9", culture);
#endif
    }

    /// <summary>A per-thread buffer of <see cref="MaximumDateTimeLength"/> characters for the writers below.</summary>
    internal static char[] Buffer => t_buffer ??= new char[MaximumDateTimeLength];

    /// <summary>
    /// Writes <see cref="DateTimeFormat"/> text into at least <see cref="MaximumDateTimeLength"/> characters without
    /// parsing the format and returns its length. Local values take the runtime's text, which looks up the offset.
    /// </summary>
    internal static int Write(DateTime value, CharBuffer destination)
    {
        if (value.Kind == DateTimeKind.Local)
        {
            string text = value.ToString(DateTimeFormat, CultureInfo.InvariantCulture);
#if NETCOREAPP
            text.AsSpan().CopyTo(destination);
#else
            text.CopyTo(0, destination, 0, text.Length);
#endif
            return text.Length;
        }

        int position = WriteDateAndTime(value.Ticks, destination);
        if (value.Kind == DateTimeKind.Utc)
        {
            destination[position++] = 'Z';
        }

        return position;
    }

    /// <summary>Writes <see cref="DateTimeOffsetFormat"/> text into at least <see cref="MaximumDateTimeLength"/> characters and returns its length.</summary>
    internal static int Write(DateTimeOffset value, CharBuffer destination)
    {
        int position = WriteDateAndTime(value.Ticks, destination);
        int offsetMinutes = (int)(value.Offset.Ticks / TimeSpan.TicksPerMinute);
        destination[position++] = offsetMinutes < 0 ? '-' : '+';
        offsetMinutes = Math.Abs(offsetMinutes);
        WriteTwoDigits(destination, position, offsetMinutes / 60);
        destination[position + 2] = ':';
        WriteTwoDigits(destination, position + 3, offsetMinutes % 60);
        return position + 5;
    }

    /// <summary>Writes yyyy-MM-ddTHH:mm:ss and, when the second has a fraction, a point and its digits without trailing zeros.</summary>
    private static int WriteDateAndTime(long ticks, CharBuffer destination)
    {
        var date = new DateTime(ticks);
#if NET8_0_OR_GREATER
        date.Deconstruct(out int year, out int month, out int day);
#else
        int year = date.Year, month = date.Month, day = date.Day;
#endif
        WriteTwoDigits(destination, 0, year / 100);
        WriteTwoDigits(destination, 2, year % 100);
        destination[4] = '-';
        WriteTwoDigits(destination, 5, month);
        destination[7] = '-';
        WriteTwoDigits(destination, 8, day);
        destination[10] = 'T';
        long dayTicks = ticks % TimeSpan.TicksPerDay;
        int seconds = (int)(dayTicks / TimeSpan.TicksPerSecond);
        WriteTwoDigits(destination, 11, seconds / 3600);
        destination[13] = ':';
        WriteTwoDigits(destination, 14, seconds / 60 % 60);
        destination[16] = ':';
        WriteTwoDigits(destination, 17, seconds % 60);
        int fraction = (int)(dayTicks % TimeSpan.TicksPerSecond);
        if (fraction == 0)
        {
            return 19;
        }

        destination[19] = '.';
        int digits = 7;
        while (fraction % 10 == 0)
        {
            fraction /= 10;
            digits--;
        }

        for (int index = 19 + digits; index > 19; index--)
        {
            destination[index] = (char)('0' + fraction % 10);
            fraction /= 10;
        }

        return 20 + digits;
    }

    private static void WriteTwoDigits(CharBuffer destination, int position, int value)
    {
        destination[position] = (char)('0' + value / 10);
        destination[position + 1] = (char)('0' + value % 10);
    }
}
