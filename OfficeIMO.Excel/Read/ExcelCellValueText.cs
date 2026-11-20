using System.Globalization;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel {
    /// <summary>
    /// The text a data reader returns for a typed cell value. Numbers use the shortest text that parses back to the
    /// same value on every target; .NET Framework's default number text keeps only 15 digits.
    /// </summary>
    internal static class ExcelCellValueText {
        private static readonly DateTime FirstExcelDate = new DateTime(1899, 12, 30);

        internal static string Format(object? value, CultureInfo culture) => value switch {
            double number => OfficeRoundTripText.Format(number, culture),
            float number => OfficeRoundTripText.Format(number, culture),
            DateTime date => FormatDate(date, culture),
            _ => Convert.ToString(value, culture) ?? string.Empty
        };

        /// <summary>
        /// Formats a cell date with the culture. A date outside the culture's calendar range, such as a time-only value
        /// before an Um Al Qura calendar's first date, cannot be formatted with that culture and uses the invariant culture.
        /// Callers that still have the cell's own text should check <see cref="FitsCalendar"/> and prefer that text.
        /// </summary>
        internal static string FormatDate(DateTime date, CultureInfo culture) =>
            FitsCalendar(date, culture) ? date.ToString(culture) : date.ToString(CultureInfo.InvariantCulture);

        /// <summary>Whether the culture's calendar can represent the date.</summary>
        internal static bool FitsCalendar(DateTime date, CultureInfo culture) {
            Calendar calendar = culture.DateTimeFormat.Calendar;
            return date >= calendar.MinSupportedDateTime && date <= calendar.MaxSupportedDateTime;
        }

        /// <summary>Whether the culture's calendar cannot represent some dates a workbook can hold.</summary>
        internal static bool CalendarLimitsCellDates(CultureInfo culture) {
            Calendar calendar = culture.DateTimeFormat.Calendar;
            return calendar.MinSupportedDateTime > FirstExcelDate || calendar.MaxSupportedDateTime < DateTime.MaxValue.Date;
        }
    }
}
