using System.Globalization;
#if NET8_0_OR_GREATER
using System.Runtime.InteropServices;
#endif
using OfficeIMO.Core.Internal;

namespace OfficeIMO.CSV;

// Values are written so that they read back unchanged: dates and times as ISO 8601 with every non-zero fraction
// digit and the kind or offset, whatever the culture, and numbers with the culture's round-trip text.
internal static partial class CsvWriter
{
    /// <summary>Formats a date with a configured format, or with the round-trip text when there is none.</summary>
    internal static string FormatDateText(DateTime value, string? format, CultureInfo culture) =>
        string.IsNullOrEmpty(format) ? OfficeRoundTripText.Format(value) : value.ToString(format, culture);

    /// <summary>Writes the round-trip text of a date and time; false when it does not fit.</summary>
    internal static bool TryFormatRoundTrip(DateTime value, Span<char> destination, out int written)
    {
        char[] buffer = OfficeRoundTripText.Buffer;
        return TryCopyRoundTrip(buffer, OfficeRoundTripText.Write(value, buffer), destination, out written);
    }

    /// <summary>Writes the round-trip text of a date, time and offset; false when it does not fit.</summary>
    internal static bool TryFormatRoundTrip(DateTimeOffset value, Span<char> destination, out int written)
    {
        char[] buffer = OfficeRoundTripText.Buffer;
        return TryCopyRoundTrip(buffer, OfficeRoundTripText.Write(value, buffer), destination, out written);
    }

    private static bool TryCopyRoundTrip(char[] buffer, int length, Span<char> destination, out int written)
    {
        if (length > destination.Length)
        {
            written = 0;
            return false;
        }

        buffer.AsSpan(0, length).CopyTo(destination);
        written = length;
        return true;
    }

#if NET8_0_OR_GREATER
    private const string DateOnlyRoundTripFormat = OfficeRoundTripText.DateOnlyFormat;
    private const string TimeOnlyRoundTripFormat = OfficeRoundTripText.TimeOnlyFormat;

    /// <summary>Formats a span-formattable value with its default text, dates and times with their round-trip text.</summary>
    internal static bool TryFormatDefault(ISpanFormattable value, Span<char> destination, out int written, CultureInfo culture) => value switch
    {
        DateTime dateTime => TryFormatRoundTrip(dateTime, destination, out written),
        DateTimeOffset dateTimeOffset => TryFormatRoundTrip(dateTimeOffset, destination, out written),
        DateOnly dateOnly => dateOnly.TryFormat(destination, out written, DateOnlyRoundTripFormat, CultureInfo.InvariantCulture),
        TimeOnly timeOnly => timeOnly.TryFormat(destination, out written, TimeOnlyRoundTripFormat, CultureInfo.InvariantCulture),
        _ => value.TryFormat(destination, out written, default, culture)
    };

    /// <inheritdoc cref="TryFormatDefault(ISpanFormattable, Span{char}, out int, CultureInfo)"/>
    internal static bool TryFormatDefault<T>(T value, Span<char> destination, out int written, CultureInfo culture)
        where T : ISpanFormattable
    {
        if (IsRoundTripDate<T>())
        {
            return TryFormatRoundTripDate(value, destination, out written);
        }

        return value.TryFormat(destination, out written, default, culture);
    }

    /// <summary>Whether <typeparamref name="T"/> is a date or time type, whose default text is the round-trip text.</summary>
    internal static bool IsRoundTripDate<T>() =>
        typeof(T) == typeof(DateTime) || typeof(T) == typeof(DateTimeOffset) || typeof(T) == typeof(DateOnly) || typeof(T) == typeof(TimeOnly);

    /// <summary>Writes the round-trip text of a date or time type as UTF-16 or UTF-8 units; the text is ASCII.</summary>
    internal static bool TryFormatRoundTripDate<T, TChar>(T value, Span<TChar> destination, out int written)
        where TChar : unmanaged
    {
        if (typeof(TChar) == typeof(char))
        {
            return TryFormatRoundTripDate(value, MemoryMarshal.Cast<TChar, char>(destination), out written);
        }

        Span<char> characters = stackalloc char[OfficeRoundTripText.MaximumDateTimeLength];
        if (!TryFormatRoundTripDate(value, characters, out written) || written > destination.Length)
        {
            written = 0;
            return false;
        }

        Span<byte> bytes = MemoryMarshal.Cast<TChar, byte>(destination);
        for (int index = 0; index < written; index++)
        {
            bytes[index] = (byte)characters[index];
        }

        return true;
    }

    private static bool TryFormatRoundTripDate<T>(T value, Span<char> destination, out int written)
    {
        if (typeof(T) == typeof(DateTime))
        {
            return TryFormatRoundTrip((DateTime)(object)value!, destination, out written);
        }

        if (typeof(T) == typeof(DateTimeOffset))
        {
            return TryFormatRoundTrip((DateTimeOffset)(object)value!, destination, out written);
        }

        if (typeof(T) == typeof(DateOnly))
        {
            return ((DateOnly)(object)value!).TryFormat(destination, out written, DateOnlyRoundTripFormat, CultureInfo.InvariantCulture);
        }

        return ((TimeOnly)(object)value!).TryFormat(destination, out written, TimeOnlyRoundTripFormat, CultureInfo.InvariantCulture);
    }
#endif
}
