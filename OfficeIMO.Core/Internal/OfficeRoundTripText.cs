using System.Globalization;
using System.Runtime.CompilerServices;
#if NETCOREAPP
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
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
        CharBuffer target = buffer;
        return new string(buffer, 0, Write(value, target, 0));
    }

    /// <summary>Formats a date, time and offset as <see cref="DateTimeOffsetFormat"/> text.</summary>
    internal static string Format(DateTimeOffset value)
    {
        char[] buffer = Buffer;
        CharBuffer target = buffer;
        return new string(buffer, 0, Write(value, target, 0));
    }

    /// <summary>
    /// Formats a date, time and offset as the runtime's "o" text: every one of the seven fraction digits and the
    /// offset, <c>+00:00</c> included.
    /// </summary>
    internal static string FormatRoundTripPattern(DateTimeOffset value)
    {
        char[] buffer = Buffer;
        CharBuffer target = buffer;
        int position = WriteDateAndTime(target, 0, value.Ticks, trimFraction: false);
        return new string(buffer, 0, WriteOffset(target, position, value.Offset.Ticks));
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

    // The writers below put the text at start in at least MaximumDateTimeLength units and return the position after
    // it: UTF-16 or UTF-8 units on .NET, characters in an array elsewhere. The ASCII text is composed eight
    // characters at a time in a 64-bit word, character i in byte i, and stored at once; a store may run past the
    // returned position, never past MaximumDateTimeLength units.

#if NETCOREAPP
    /// <summary>Writes <see cref="DateTimeFormat"/> text; a local value takes the offset the runtime's <c>K</c> writes.</summary>
    internal static int Write<TChar>(DateTime value, Span<TChar> destination, int start) where TChar : unmanaged
#else
    /// <summary>Writes <see cref="DateTimeFormat"/> text; a local value takes the offset the runtime's <c>K</c> writes.</summary>
    internal static int Write(DateTime value, char[] destination, int start)
#endif
    {
        EnsureRoom(destination, start);
        int position = WriteDateAndTime(destination, start, value.Ticks, trimFraction: true);
        switch (value.Kind)
        {
            case DateTimeKind.Utc:
                destination[position] = Unit(destination, 'Z');
                return position + 1;
            case DateTimeKind.Local:
                return WriteOffset(destination, position, TimeZoneInfo.Local.GetUtcOffset(value).Ticks);
            default:
                return position;
        }
    }

#if NETCOREAPP
    /// <summary>Writes <see cref="DateTimeOffsetFormat"/> text.</summary>
    internal static int Write<TChar>(DateTimeOffset value, Span<TChar> destination, int start) where TChar : unmanaged
#else
    /// <summary>Writes <see cref="DateTimeOffsetFormat"/> text.</summary>
    internal static int Write(DateTimeOffset value, char[] destination, int start)
#endif
    {
        EnsureRoom(destination, start);
        int position = WriteDateAndTime(destination, start, value.Ticks, trimFraction: true);
        return WriteOffset(destination, position, value.Offset.Ticks);
    }

    /// <summary>
    /// Writes yyyy-MM-ddTHH:mm:ss and the point and fraction digits: all seven, or, when trimmed, none for a whole
    /// second and no trailing zeros otherwise.
    /// </summary>
#if NETCOREAPP
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static int WriteDateAndTime<TChar>(Span<TChar> destination, int start, long ticks, bool trimFraction) where TChar : unmanaged
#else
    private static int WriteDateAndTime(char[] destination, int start, long ticks, bool trimFraction)
#endif
    {
        // Ticks of the day fit in 40 bits, so the clock is split with unsigned constant divisions.
        ulong dayNumber = (ulong)ticks / TimeSpan.TicksPerDay;
        ulong dayTicks = (ulong)ticks - dayNumber * TimeSpan.TicksPerDay;

        // Rows often repeat a day, so the text of the last one is kept: "yyyy-MM-" and "dd".
        ref DayText last = ref t_lastDay;
        if (last.Number != dayNumber + 1)
        {
            var date = new DateTime(ticks);
#if NET8_0_OR_GREATER
            date.Deconstruct(out int year, out int month, out int day);
#else
            int year = date.Year, month = date.Month, day = date.Day;
#endif
            uint century = (uint)year / 100;
            last.YearAndMonth = TwoDigits(century) | (ulong)TwoDigits((uint)year - century * 100) << 16 | (ulong)'-' << 32 |
                (ulong)TwoDigits((uint)month) << 40 | (ulong)'-' << 56;
            last.Day = TwoDigits((uint)day) | (uint)'T' << 16;
            last.Number = dayNumber + 1;
        }

        Store8(destination, start, last.YearAndMonth);
        if (dayTicks == 0)
        {
            // "ddT00:00" and ":00".
            Store8(destination, start + 8, last.Day | MidnightClock);
            Store4(destination, start + 16, ':' | (uint)'0' << 8 | (uint)'0' << 16);
            if (trimFraction)
            {
                return start + 19;
            }
        }

        uint seconds = (uint)(dayTicks / TimeSpan.TicksPerSecond);
        uint fraction = (uint)(dayTicks - seconds * (ulong)TimeSpan.TicksPerSecond);
        uint minutes = seconds / 60;
        uint hours = minutes / 60;

        // "ddTHH:mm".
        Store8(destination, start + 8, last.Day | (ulong)TwoDigits(hours) << 24 | (ulong)':' << 40 |
            (ulong)TwoDigits(minutes - hours * 60) << 48);
        ulong secondText = ':' | (ulong)TwoDigits(seconds - minutes * 60) << 8;
        if (trimFraction && fraction == 0)
        {
            Store4(destination, start + 16, (uint)secondText);
            return start + 19;
        }

        // ":ss.ffff" and "fff".
        uint high = fraction / 1_000;
        uint low = fraction - high * 1_000;
        uint lowHundreds = low / 100;
        Store8(destination, start + 16, secondText | (ulong)'.' << 24 | (ulong)TwoDigits(high / 100) << 32 |
            (ulong)TwoDigits(high % 100) << 48);
        Store4(destination, start + 24, '0' + lowHundreds | TwoDigits(low - lowHundreds * 100) << 8);
        if (!trimFraction)
        {
            return start + 27;
        }

        int end = start + 27;
        while (fraction % 10 == 0)
        {
            fraction /= 10;
            end--;
        }

        return end;
    }

    /// <summary>Writes an offset as +hh:mm or -hh:mm, as the runtime's <c>zzz</c> and <c>K</c> do.</summary>
#if NETCOREAPP
    private static int WriteOffset<TChar>(Span<TChar> destination, int position, long offsetTicks) where TChar : unmanaged
#else
    private static int WriteOffset(char[] destination, int position, long offsetTicks)
#endif
    {
        int offsetMinutes = (int)(offsetTicks / TimeSpan.TicksPerMinute);
        uint sign = offsetMinutes < 0 ? '-' : '+';
        uint magnitude = (uint)Math.Abs(offsetMinutes);
        uint offsetHours = magnitude / 60;
        Store4(destination, position, sign | TwoDigits(offsetHours) << 8 | (uint)':' << 24);
        uint minuteText = TwoDigits(magnitude - offsetHours * 60);
        destination[position + 4] = Unit(destination, (char)(minuteText & 0xFF));
        destination[position + 5] = Unit(destination, (char)(minuteText >> 8));
        return position + 6;
    }

    /// <summary>"00:00" after "ddT", in the bytes of the second eight characters.</summary>
    private const ulong MidnightClock = (ulong)'0' << 24 | (ulong)'0' << 32 | (ulong)':' << 40 | (ulong)'0' << 48 | (ulong)'0' << 56;

    /// <summary>The text of a day: its number plus one (zero when none was written on this thread), "yyyy-MM-", and "dd" and 'T'.</summary>
    private struct DayText
    {
        public ulong Number;
        public ulong YearAndMonth;
        public uint Day;
    }

    [ThreadStatic]
    private static DayText t_lastDay;

    /// <summary>The two ASCII digits of a value below 100, the first in the low byte.</summary>
    private static uint TwoDigits(uint value)
    {
        uint tens = value / 10;
        return '0' + tens | ('0' + value - tens * 10) << 8;
    }

#if NETCOREAPP
    private static void EnsureRoom<TChar>(Span<TChar> destination, int start) where TChar : unmanaged
    {
        if ((uint)start > (uint)destination.Length || destination.Length - start < MaximumDateTimeLength)
        {
            throw new ArgumentException("The destination is too short for the date text.", nameof(destination));
        }
    }

    /// <summary>The unit of an ASCII character; the span only selects the unit type.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static TChar Unit<TChar>(Span<TChar> _, char character) where TChar : unmanaged =>
        typeof(TChar) == typeof(byte) ? ByteUnit<TChar>((byte)character) : Unsafe.As<char, TChar>(ref character);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static TChar ByteUnit<TChar>(byte value) where TChar : unmanaged => Unsafe.As<byte, TChar>(ref value);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void Store8<TChar>(Span<TChar> destination, int position, ulong text) where TChar : unmanaged
    {
        ref TChar target = ref Unsafe.Add(ref MemoryMarshal.GetReference(destination), position);
        if (!BitConverter.IsLittleEndian)
        {
            for (int index = 0; index < 8; index++)
            {
                Unsafe.Add(ref target, index) = Unit(destination, (char)(byte)(text >> (index * 8)));
            }
        }
        else if (typeof(TChar) == typeof(byte))
        {
            Unsafe.WriteUnaligned(ref Unsafe.As<TChar, byte>(ref target), text);
        }
        else
        {
            Vector128.WidenLower(Vector128.CreateScalarUnsafe(text).AsByte())
                .StoreUnsafe(ref Unsafe.As<TChar, ushort>(ref target));
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void Store4<TChar>(Span<TChar> destination, int position, uint text) where TChar : unmanaged
    {
        ref TChar target = ref Unsafe.Add(ref MemoryMarshal.GetReference(destination), position);
        if (!BitConverter.IsLittleEndian)
        {
            for (int index = 0; index < 4; index++)
            {
                Unsafe.Add(ref target, index) = Unit(destination, (char)(byte)(text >> (index * 8)));
            }
        }
        else if (typeof(TChar) == typeof(byte))
        {
            Unsafe.WriteUnaligned(ref Unsafe.As<TChar, byte>(ref target), text);
        }
        else
        {
            ulong wide = (text & 0xFF) | (ulong)(text & 0xFF00) << 8 | (ulong)(text & 0xFF0000) << 16 | (ulong)(text & 0xFF000000) << 24;
            Unsafe.WriteUnaligned(ref Unsafe.As<TChar, byte>(ref target), wide);
        }
    }
#else
    private static void EnsureRoom(char[] destination, int start)
    {
        if ((uint)start > (uint)destination.Length || destination.Length - start < MaximumDateTimeLength)
        {
            throw new ArgumentException("The destination is too short for the date text.", nameof(destination));
        }
    }

    private static char Unit(char[] _, char character) => character;

    private static void Store8(char[] destination, int position, ulong text)
    {
        for (int index = 0; index < 8; index++)
        {
            destination[position + index] = (char)(byte)(text >> (index * 8));
        }
    }

    private static void Store4(char[] destination, int position, uint text)
    {
        for (int index = 0; index < 4; index++)
        {
            destination[position + index] = (char)(byte)(text >> (index * 8));
        }
    }
#endif
}