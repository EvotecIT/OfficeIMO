using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Parses the ISO 8601 text of a date cell (<c>t="d"</c>). ISO text is parsed with the invariant culture, so a
    /// culture's calendar never changes the stored date. Text with only a time of day is a time-only value: it reads
    /// as the time on the OLE Automation base date, December 30, 1899, as a time-only serial does, and converts to
    /// the serial fraction of the day, never to a time on today's date. An offset on such text is ignored because
    /// there is no date to apply it to.
    /// </summary>
    internal static class ExcelIsoDateText {
        private static readonly DateTime TimeOnlyBase = new DateTime(1899, 12, 30);

        private static readonly string[] TimeOfDayFormats = { "HH:mm:ss.FFFFFFF", "HH:mm:ss", "HH:mm" };

        /// <summary>
        /// Parses date cell text. ISO text is parsed invariantly; other text keeps the reader's culture.
        /// </summary>
        internal static bool TryParse(string? text, CultureInfo culture, DateTimeStyles styles, out DateTime value) {
            if (TryParseTimeOfDay(text, out TimeSpan time)) {
                value = TimeOnlyBase + time;
                return true;
            }

            if (text == null) {
                value = default;
                return false;
            }

            if (LooksLikeIsoDate(text) && DateTime.TryParse(text, CultureInfo.InvariantCulture, styles, out value)) {
                return true;
            }

            return DateTime.TryParse(text, culture, styles, out value);
        }

        /// <summary>Converts date cell text to a serial in the workbook's date system; a time of day is its fraction.</summary>
        internal static bool TryGetSerial(string? text, ExcelDateSystem dateSystem, DateTimeStyles styles, out double serial) {
            if (TryParseTimeOfDay(text, out TimeSpan time)) {
                serial = time.TotalDays;
                return true;
            }

            if (text != null && DateTime.TryParse(text, CultureInfo.InvariantCulture, styles, out DateTime date)) {
                serial = ExcelDateSystemConverter.ToSerial(date, dateSystem);
                return true;
            }

            serial = 0D;
            return false;
        }

        /// <summary>
        /// Resolves date cell text for display: the date to format and its display serial. A time of day displays as
        /// the same fraction stored as a number does.
        /// </summary>
        internal static bool TryParseForDisplay(string? text, ExcelDateSystem dateSystem, out DateTime date, out double serial) {
            if (TryParseTimeOfDay(text, out TimeSpan time)) {
                serial = time.TotalDays;
                date = ExcelDateSystemConverter.FromSerialForDisplay(serial, dateSystem);
                return true;
            }

            if (text != null && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out date)) {
                serial = ExcelDateSystemConverter.ToSerialForDisplay(date, dateSystem);
                return true;
            }

            date = default;
            serial = 0D;
            return false;
        }

        private static bool LooksLikeIsoDate(string text) {
            int start = 0;
            while (start < text.Length && char.IsWhiteSpace(text[start])) start++;
            if (start < text.Length && (text[start] == '-' || text[start] == '+')) start++;
            return text.Length - start >= 5
                && IsDigit(text[start]) && IsDigit(text[start + 1]) && IsDigit(text[start + 2]) && IsDigit(text[start + 3])
                && text[start + 4] == '-';
        }

        /// <summary>Parses ISO 8601 time-of-day text, <c>hh:mm[:ss[.fff]]</c> with an optional <c>T</c> prefix and offset.</summary>
        private static bool TryParseTimeOfDay(string? text, out TimeSpan time) {
            time = default;
            if (text == null) return false;
            string trimmed = text.Trim();
            int start = trimmed.Length > 0 && (trimmed[0] == 'T' || trimmed[0] == 't') ? 1 : 0;
            if (trimmed.Length - start < 5 || !IsDigit(trimmed[start]) || !IsDigit(trimmed[start + 1]) || trimmed[start + 2] != ':') {
                return false;
            }

            int end = trimmed.Length;
            if (trimmed[end - 1] == 'Z' || trimmed[end - 1] == 'z') {
                end--;
            } else {
                int sign = trimmed.LastIndexOfAny(new[] { '+', '-' });
                if (sign > start) end = sign;
            }

            if (!DateTime.TryParseExact(trimmed.Substring(start, end - start), TimeOfDayFormats, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out DateTime clock)) {
                return false;
            }

            time = clock.TimeOfDay;
            return true;
        }

        private static bool IsDigit(char value) => value >= '0' && value <= '9';
    }
}
