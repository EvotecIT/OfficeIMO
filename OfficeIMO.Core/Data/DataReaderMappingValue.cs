using System.Data.Common;

namespace OfficeIMO.Data;

/// <summary>Preserves a provider's original value and numeric getter result across conversion and row snapshots.</summary>
internal sealed class DataReaderMappingValue {
    private DataReaderMappingValue(object original, double numeric) {
        Original = original;
        Numeric = numeric;
    }

    internal object Original { get; }
    internal double Numeric { get; }

    internal static object? Read(DbDataReader reader, int ordinal, Type targetType) {
        object? value = reader.GetValue(ordinal);
        Type effective = Nullable.GetUnderlyingType(targetType) ?? targetType;
        // A number read as a date takes the reader's own date interpretation, such as a workbook's date system.
        if ((value is double || value is decimal) && IsDate(effective)) return ReadDate(reader, ordinal, value);
        // A duration is the day count of a date value's serial, not its calendar date.
        if (value is not DateTime || !(IsNumeric(effective) || effective == typeof(TimeSpan))) return value;
        try {
            return new DataReaderMappingValue(value, reader.GetDouble(ordinal));
        } catch (InvalidCastException) {
        } catch (FormatException) {
        } catch (OverflowException) {
        } catch (NotSupportedException) {
        } catch (NotImplementedException) {
        }
        return value;
    }

    /// <summary>
    /// Converts a day serial to that many days, rounded to the nearest millisecond. Excel object mapping and typed
    /// readers share this rule so every route reads the same duration.
    /// </summary>
    internal static bool TryGetDuration(double days, out TimeSpan duration) {
        double milliseconds = Math.Round(days * 86_400_000d, MidpointRounding.AwayFromZero);
        // Also rejects NaN and infinities.
        if (!(Math.Abs(milliseconds) <= TimeSpan.MaxValue.Ticks / TimeSpan.TicksPerMillisecond)) {
            duration = default;
            return false;
        }

        duration = new TimeSpan((long)milliseconds * TimeSpan.TicksPerMillisecond);
        return true;
    }

    private static bool IsDate(Type type) => type == typeof(DateTime) || type == typeof(DateTimeOffset)
#if NET6_0_OR_GREATER
        || type == typeof(DateOnly)
#endif
        ;

    private static object ReadDate(DbDataReader reader, int ordinal, object value) {
        try {
            return reader.GetDateTime(ordinal);
        } catch (InvalidCastException) {
        } catch (FormatException) {
        } catch (OverflowException) {
        } catch (ArgumentException) {
        } catch (NotSupportedException) {
        }
        return value;
    }

    internal static bool IsNumeric(Type type) => type == typeof(double) || type == typeof(float)
        || type == typeof(decimal) || type == typeof(int) || type == typeof(long)
        || type == typeof(short) || type == typeof(byte) || type == typeof(sbyte)
        || type == typeof(ushort) || type == typeof(uint) || type == typeof(ulong);
}
