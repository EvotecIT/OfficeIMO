using System.Globalization;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    /// <summary>Formats a date with a configured format; simple numeric templates skip the runtime's format parsing.</summary>
    internal static bool TryFormatDate(DateTime value, Span<char> destination, out int written, string? format, CultureInfo culture) =>
        format is not null && CsvDateTemplate.TryFormat(value, format, culture, destination, out written)
#if NET8_0_OR_GREATER
        || value.TryFormat(destination, out written, format, culture);
#else
        || CopyFormatted(value.ToString(format, culture), destination, out written);

    /// <summary>Copies formatted text when it fits; false leaves the value to the caller's string path.</summary>
    internal static bool CopyFormatted(string text, Span<char> destination, out int written)
    {
        written = 0;
        if (text.Length > destination.Length)
        {
            return false;
        }

        text.AsSpan().CopyTo(destination);
        written = text.Length;
        return true;
    }
#endif
}

/// <summary>
/// A custom date format made only of zero-padded numeric fields (<c>yyyy</c>, <c>MM</c>, <c>dd</c>, <c>HH</c>,
/// <c>mm</c>, <c>ss</c>, <c>f</c> to <c>fffffff</c>) and separator characters, written for the invariant culture
/// without parsing the format again for every value. The text is what <see cref="DateTime.ToString(string, IFormatProvider)"/>
/// writes; any other format or culture is left to the runtime.
/// </summary>
internal sealed class CsvDateTemplate
{
    private const byte Literal = 0, Year = 1, Month = 2, Day = 3, Hour = 4, Minute = 5, Second = 6, Fraction = 7;
    private static readonly int[] FractionDivisors = { 0, 1_000_000, 100_000, 10_000, 1_000, 100, 10, 1 };

    /// <summary>The last format seen; writers keep one format, so one entry serves them.</summary>
    private static CsvDateTemplate? s_last;

    private readonly string _format;
    private readonly (byte Kind, char Value)[]? _parts;
    private readonly int _length;

    private CsvDateTemplate(string format)
    {
        _format = format;
        _parts = Parse(format, out _length);
    }

#if !NET8_0_OR_GREATER
    /// <summary>The invariant culture's general date and time text: its short date and long time patterns.</summary>
    private static readonly CsvDateTemplate s_invariantGeneral = new("MM/dd/yyyy HH:mm:ss");

    /// <summary>Writes what <c>value.ToString(CultureInfo.InvariantCulture)</c> returns.</summary>
    internal static bool TryFormatInvariantGeneral(DateTime value, Span<char> destination, out int written) =>
        s_invariantGeneral.TryWrite(value, destination, out written);
#endif

    /// <summary>Writes <paramref name="value"/> when the format is a supported template; false leaves it to the runtime.</summary>
    internal static bool TryFormat(DateTime value, string format, CultureInfo culture, Span<char> destination, out int written)
    {
        written = 0;
        if (!ReferenceEquals(culture, CultureInfo.InvariantCulture))
        {
            return false;
        }

        CsvDateTemplate? template = s_last;
        if (template is null || !string.Equals(template._format, format, StringComparison.Ordinal))
        {
            s_last = template = new CsvDateTemplate(format);
        }

        return template.TryWrite(value, destination, out written);
    }

    private bool TryWrite(DateTime value, Span<char> destination, out int written)
    {
        written = 0;
        if (_parts is null || destination.Length < _length)
        {
            return false;
        }

#if NET8_0_OR_GREATER
        value.Deconstruct(out int year, out int month, out int day);
#else
        GetDate(value.Ticks, out int year, out int month, out int day);
#endif
        long dayTicks = value.Ticks % TimeSpan.TicksPerDay;
        int seconds = (int)(dayTicks / TimeSpan.TicksPerSecond);
        int fraction = (int)(dayTicks % TimeSpan.TicksPerSecond);
        int position = 0;
        foreach ((byte kind, char literal) in _parts)
        {
            switch (kind)
            {
                case Year:
                    WriteTwoDigits(destination, position, year / 100);
                    WriteTwoDigits(destination, position + 2, year % 100);
                    position += 4;
                    break;
                case Month:
                    WriteTwoDigits(destination, position, month);
                    position += 2;
                    break;
                case Day:
                    WriteTwoDigits(destination, position, day);
                    position += 2;
                    break;
                case Hour:
                    WriteTwoDigits(destination, position, seconds / 3600);
                    position += 2;
                    break;
                case Minute:
                    WriteTwoDigits(destination, position, seconds / 60 % 60);
                    position += 2;
                    break;
                case Second:
                    WriteTwoDigits(destination, position, seconds % 60);
                    position += 2;
                    break;
                case Fraction:
                    int digits = literal;
                    int truncated = fraction / FractionDivisors[digits];
                    for (int index = position + digits - 1; index >= position; index--)
                    {
                        destination[index] = (char)('0' + truncated % 10);
                        truncated /= 10;
                    }

                    position += digits;
                    break;
                default:
                    destination[position++] = literal;
                    break;
            }
        }

        written = position;
        return true;
    }

#if !NET8_0_OR_GREATER
    private static readonly int[] DaysToMonth365 = { 0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365 };
    private static readonly int[] DaysToMonth366 = { 0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366 };

    /// <summary>
    /// The Gregorian year, month and day of a tick count in one pass, as <c>DateTime.Deconstruct</c> gives them on
    /// newer runtimes; the separate Year, Month and Day properties each repeat the whole calculation.
    /// </summary>
    private static void GetDate(long ticks, out int year, out int month, out int day)
    {
        int days = (int)(ticks / TimeSpan.TicksPerDay);
        int centuries400 = days / 146_097;
        days -= centuries400 * 146_097;
        int centuries = Math.Min(days / 36_524, 3);
        days -= centuries * 36_524;
        int years4 = days / 1_461;
        days -= years4 * 1_461;
        int years = Math.Min(days / 365, 3);
        days -= years * 365;
        year = centuries400 * 400 + centuries * 100 + years4 * 4 + years + 1;
        int[] monthStarts = years == 3 && (years4 != 24 || centuries == 3) ? DaysToMonth366 : DaysToMonth365;
        month = (days >> 5) + 1;
        while (days >= monthStarts[month])
        {
            month++;
        }

        day = days - monthStarts[month - 1] + 1;
    }
#endif

    private static void WriteTwoDigits(Span<char> destination, int position, int value)
    {
        destination[position] = (char)('0' + value / 10);
        destination[position + 1] = (char)('0' + value % 10);
    }

    /// <summary>
    /// Splits the format into parts, or returns null for anything else: a one-character format is a standard
    /// format, and other specifiers, quotes and escapes keep their runtime meaning. The invariant culture's
    /// time and date separators are ':' and '/'.
    /// </summary>
    private static (byte Kind, char Value)[]? Parse(string format, out int length)
    {
        length = 0;
        if (format.Length < 2)
        {
            return null;
        }

        var parts = new List<(byte Kind, char Value)>();
        for (int index = 0; index < format.Length;)
        {
            char current = format[index];
            if (current is '\'' or '"')
            {
                // Quoted text is copied as it is; a backslash inside it or a missing closing quote is left to the runtime.
                int close = format.IndexOf(current, index + 1);
                if (close < 0 || format.IndexOf('\\', index + 1, close - index - 1) >= 0)
                {
                    return null;
                }

                for (int literal = index + 1; literal < close; literal++)
                {
                    parts.Add((Literal, format[literal]));
                }

                length += close - index - 1;
                index = close + 1;
                continue;
            }

            if (current == '\\')
            {
                if (index + 1 >= format.Length)
                {
                    return null;
                }

                parts.Add((Literal, format[index + 1]));
                length++;
                index += 2;
                continue;
            }

            int run = 1;
            while (index + run < format.Length && format[index + run] == current)
            {
                run++;
            }

            (byte Kind, char Value) part;
            int width;
            switch (current)
            {
                case 'y' when run == 4: part = (Year, '\0'); width = 4; break;
                case 'M' when run == 2: part = (Month, '\0'); width = 2; break;
                case 'd' when run == 2: part = (Day, '\0'); width = 2; break;
                case 'H' when run == 2: part = (Hour, '\0'); width = 2; break;
                case 'm' when run == 2: part = (Minute, '\0'); width = 2; break;
                case 's' when run == 2: part = (Second, '\0'); width = 2; break;
                case 'f' when run <= 7: part = (Fraction, (char)run); width = run; break;
                case '-' or 'T' or ' ' or '.' or '_' or ':' or '/': part = (Literal, current); width = 1; run = 1; break;
                default: return null;
            }

            parts.Add(part);
            length += width;
            index += run;
        }

        return parts.ToArray();
    }
}
