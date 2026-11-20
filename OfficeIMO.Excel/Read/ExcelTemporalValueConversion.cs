using System.Globalization;
using OfficeIMO.Data;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Converts cell values to date and time destinations for object mapping. A number is a serial in the workbook's
    /// date system; a duration is the serial's day count, and a time of day is the serial's fraction. A value that has
    /// no such meaning is an error rather than the destination's default.
    /// </summary>
    internal static class ExcelTemporalValueConversion {
        private static readonly DateTime OleAutomationEpoch = new DateTime(1899, 12, 30);

        /// <summary>Whether <paramref name="destinationType"/> is converted here instead of the general converter.</summary>
        internal static bool Handles(Type destinationType) =>
            destinationType == typeof(TimeSpan)
            || destinationType == typeof(DateTimeOffset)
#if NET6_0_OR_GREATER
            || destinationType == typeof(DateOnly)
            || destinationType == typeof(TimeOnly)
#endif
            || destinationType == typeof(DateTime);

        /// <summary>
        /// Converts <paramref name="value"/> to <paramref name="destinationType"/>, returns null for empty text, or throws
        /// <see cref="DataMappingException"/> naming the destination when the value is not a date, time or duration.
        /// </summary>
        /// <param name="value">The cell value: a number, date, duration or text.</param>
        /// <param name="destinationType">A type for which <see cref="Handles"/> is true.</param>
        /// <param name="dateSystem">The workbook's date system, applied to numbers read as calendar dates.</param>
        /// <param name="culture">The culture for text that is not ISO 8601.</param>
        /// <param name="target">The destination member, used in the error message.</param>
        /// <param name="errorValuePolicy">Whether the error message may include the value.</param>
        internal static object? Convert(object value, Type destinationType, ExcelDateSystem dateSystem, CultureInfo culture,
            string target, DataMappingErrorValuePolicy errorValuePolicy) {
            // Empty text is a blank cell, which leaves the destination unset as other blank cells do.
            if (value is string blank && string.IsNullOrWhiteSpace(blank)) return null;
            object? result;
            try {
                result = TryConvert(value, destinationType, dateSystem, culture);
            } catch (ArgumentException) {
                result = null;
            } catch (OverflowException) {
                result = null;
            }

            if (result != null) return result;
            string source = errorValuePolicy == DataMappingErrorValuePolicy.Redact
                ? "The cell value"
                : value is double number
                    ? "The number " + number.ToString("R", CultureInfo.InvariantCulture)
                    : $"The {value.GetType().Name} value '{System.Convert.ToString(value, CultureInfo.InvariantCulture)}'";
            throw new DataMappingException($"{source} cannot be assigned to {target}: it is not a valid {destinationType.Name}.");
        }

        private static object? TryConvert(object value, Type destinationType, ExcelDateSystem dateSystem, CultureInfo culture) {
            if (destinationType.IsInstanceOfType(value)) return value;
            if (destinationType == typeof(TimeSpan)) {
                return value switch {
                    double serial when DataReaderMappingValue.TryGetDuration(serial, out TimeSpan duration) => duration,
                    DateTime date => date - OleAutomationEpoch,
#if NET6_0_OR_GREATER
                    TimeOnly time => time.ToTimeSpan(),
#endif
                    string text when TimeSpan.TryParseExact(text.Trim(), "c", CultureInfo.InvariantCulture, out TimeSpan invariant) => invariant,
                    string text when TimeSpan.TryParse(text, culture, out TimeSpan parsed) => parsed,
                    _ => null
                };
            }

            if (destinationType == typeof(DateTime)) {
                return value switch {
                    double serial => ExcelDateSystemConverter.FromSerial(serial, dateSystem),
                    DateTimeOffset offset => offset.DateTime,
#if NET6_0_OR_GREATER
                    DateOnly date => date.ToDateTime(TimeOnly.MinValue),
#endif
                    string text => ParseDateTime(text, culture),
                    _ => null
                };
            }

            if (destinationType == typeof(DateTimeOffset)) {
                return value switch {
                    // A cell date has no offset; like DateTimeOffset(DateTime), it takes the local offset for that date.
                    double serial => new DateTimeOffset(ExcelDateSystemConverter.FromSerial(serial, dateSystem)),
                    DateTime date => new DateTimeOffset(date),
                    string text when DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTimeOffset invariant) => invariant,
                    string text when DateTimeOffset.TryParse(text, culture, DateTimeStyles.None, out DateTimeOffset parsed) => parsed,
                    _ => null
                };
            }

#if NET6_0_OR_GREATER
            if (destinationType == typeof(DateOnly)) {
                return value switch {
                    double serial => DateOnly.FromDateTime(ExcelDateSystemConverter.FromSerial(serial, dateSystem)),
                    DateTime date => DateOnly.FromDateTime(date),
                    DateTimeOffset offset => DateOnly.FromDateTime(offset.DateTime),
                    string text when ParseDateTime(text, culture) is DateTime parsed => DateOnly.FromDateTime(parsed),
                    _ => null
                };
            }

            if (destinationType == typeof(TimeOnly)) {
                return value switch {
                    double serial => TimeOnly.FromDateTime(ExcelDateSystemConverter.FromElapsedSerial(serial)),
                    DateTime date => TimeOnly.FromDateTime(date),
                    TimeSpan duration when duration >= TimeSpan.Zero && duration < TimeSpan.FromDays(1) => TimeOnly.FromTimeSpan(duration),
                    string text when TimeOnly.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out TimeOnly invariant) => invariant,
                    string text when TimeOnly.TryParse(text, culture, DateTimeStyles.None, out TimeOnly parsed) => parsed,
                    _ => null
                };
            }
#endif
            return null;
        }

        private static object? ParseDateTime(string text, CultureInfo culture) =>
            DateTime.TryParse(text, culture, DateTimeStyles.AssumeLocal, out DateTime parsed) ? parsed : null;
    }
}
