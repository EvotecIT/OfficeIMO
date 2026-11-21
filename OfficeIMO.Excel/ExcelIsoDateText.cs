using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Parses the ISO 8601 text of a date cell (<c>t="d"</c>). ISO text is parsed with the invariant culture, so a
    /// culture's calendar never changes the stored date. Text with only a time of day is a time-only value: it reads
    /// as the time on the OLE Automation base date, December 30, 1899, as a time-only serial does, and converts to
    /// the serial fraction of the day, never to a time on today's date. A valid offset on such text is ignored because
    /// there is no date to apply it to; time-shaped text that is not a valid ISO time of day is rejected.
    /// </summary>
    internal static class ExcelIsoDateText {
        private static readonly DateTime TimeOnlyBase = new DateTime(1899, 12, 30);

        /// <summary>
        /// Parses date cell text. ISO text is parsed invariantly; other text keeps the reader's culture.
        /// </summary>
        internal static bool TryParse(string? text, CultureInfo culture, DateTimeStyles styles, out DateTime value) {
            if (TryParseTimeOfDay(text, out TimeSpan time, out bool isTimeOfDay)) {
                value = TimeOnlyBase + time;
                return true;
            }

            if (isTimeOfDay || text == null) {
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
            if (TryParseTimeOfDay(text, out TimeSpan time, out bool isTimeOfDay)) {
                serial = ExcelDateSystemConverter.ToSerial(time);
                return true;
            }

            if (!isTimeOfDay && text != null && DateTime.TryParse(text, CultureInfo.InvariantCulture, styles, out DateTime date)) {
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
            if (TryParseTimeOfDay(text, out TimeSpan time, out bool isTimeOfDay)) {
                serial = time.TotalDays;
                date = ExcelDateSystemConverter.FromSerialForDisplay(serial, dateSystem);
                return true;
            }

            if (!isTimeOfDay && text != null && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out date)) {
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

        /// <summary>
        /// Parses ISO 8601 time-of-day text, <c>hh:mm[:ss[.fraction]]</c> with an optional <c>T</c> prefix and an optional
        /// <c>Z</c> or <c>±hh[:mm]</c> offset. <paramref name="isTimeOfDay"/> reports text that starts like a time of day,
        /// so a caller rejects malformed time text instead of handing it to a general parser that would put it on
        /// today's date. A fraction longer than the seven digits a tick holds is truncated.
        /// </summary>
        private static bool TryParseTimeOfDay(string? text, out TimeSpan time, out bool isTimeOfDay) {
            time = default;
            isTimeOfDay = false;
            if (text == null) return false;
            string value = text.Trim();
            int index = value.Length > 0 && (value[0] == 'T' || value[0] == 't') ? 1 : 0;
            if (value.Length - index < 3 || !IsDigit(value[index]) || !IsDigit(value[index + 1]) || value[index + 2] != ':') {
                return false;
            }

            isTimeOfDay = true;
            if (!TryReadTwoDigits(value, ref index, 23, out int hours) || !TryRead(value, ref index, ':')
                || !TryReadTwoDigits(value, ref index, 59, out int minutes)) {
                return false;
            }

            int seconds = 0;
            long fractionTicks = 0;
            if (TryRead(value, ref index, ':')) {
                if (!TryReadTwoDigits(value, ref index, 59, out seconds)) return false;
                if (index < value.Length && (value[index] == '.' || value[index] == ',')) {
                    index++;
                    int digits = 0;
                    while (index < value.Length && IsDigit(value[index])) {
                        if (digits < 7) fractionTicks = (fractionTicks * 10) + (value[index] - '0');
                        digits++;
                        index++;
                    }
                    if (digits == 0) return false;
                    for (int padding = digits; padding < 7; padding++) fractionTicks *= 10;
                }
            }

            if (!IsValidOffset(value, index)) return false;

            time = new TimeSpan(hours, minutes, seconds) + TimeSpan.FromTicks(fractionTicks);
            return true;
        }

        // Accepts the end of the text, Z, or ±hh, ±hhmm or ±hh:mm with an offset of at most 14 hours.
        private static bool IsValidOffset(string value, int index) {
            if (index == value.Length) return true;
            if (value[index] == 'Z' || value[index] == 'z') return index + 1 == value.Length;
            if (value[index] != '+' && value[index] != '-') return false;
            index++;
            if (!TryReadTwoDigits(value, ref index, 14, out int hours)) return false;
            if (index == value.Length) return true;
            TryRead(value, ref index, ':');
            return TryReadTwoDigits(value, ref index, 59, out int minutes)
                && index == value.Length
                && (hours < 14 || minutes == 0);
        }

        private static bool TryReadTwoDigits(string value, ref int index, int maximum, out int number) {
            number = 0;
            if (index + 2 > value.Length || !IsDigit(value[index]) || !IsDigit(value[index + 1])) return false;
            number = ((value[index] - '0') * 10) + (value[index + 1] - '0');
            index += 2;
            return number <= maximum;
        }

        private static bool TryRead(string value, ref int index, char expected) {
            if (index >= value.Length || value[index] != expected) return false;
            index++;
            return true;
        }

        private static bool IsDigit(char value) => value >= '0' && value <= '9';
    }
}
