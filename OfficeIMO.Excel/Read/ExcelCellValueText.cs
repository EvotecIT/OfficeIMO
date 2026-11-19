using System.Globalization;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel {
    /// <summary>
    /// The text a data reader returns for a typed cell value. Numbers use the shortest text that parses back to the
    /// same value on every target; .NET Framework's default number text keeps only 15 digits.
    /// </summary>
    internal static class ExcelCellValueText {
        internal static string Format(object? value, CultureInfo culture) => value switch {
            double number => OfficeRoundTripText.Format(number, culture),
            float number => OfficeRoundTripText.Format(number, culture),
            // A date outside the culture's calendar range, such as a time-only value under an Um Al Qura calendar,
            // cannot be formatted with that culture.
            DateTime date when date < culture.DateTimeFormat.Calendar.MinSupportedDateTime
                || date > culture.DateTimeFormat.Calendar.MaxSupportedDateTime => date.ToString(CultureInfo.InvariantCulture),
            _ => Convert.ToString(value, culture) ?? string.Empty
        };
    }
}
