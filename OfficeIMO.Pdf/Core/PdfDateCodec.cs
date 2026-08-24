using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfDateCodec {
    internal static string Format(DateTimeOffset value) {
        string sign = value.Offset < TimeSpan.Zero ? "-" : "+";
        TimeSpan offset = value.Offset.Duration();
        return "D:" + value.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture)
            + sign
            + offset.Hours.ToString("00", CultureInfo.InvariantCulture)
            + "'"
            + offset.Minutes.ToString("00", CultureInfo.InvariantCulture)
            + "'";
    }

    internal static DateTimeOffset? TryParse(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        string raw = value!.StartsWith("D:", StringComparison.Ordinal) ? value.Substring(2) : value;
        if (raw.Length < 4 || !TryPart(raw, 0, 4, out int year)) return null;
        if (!TryOptionalPart(raw, 4, out int parsedMonth) ||
            !TryOptionalPart(raw, 6, out int parsedDay) ||
            !TryOptionalPart(raw, 8, out int parsedHour) ||
            !TryOptionalPart(raw, 10, out int parsedMinute) ||
            !TryOptionalPart(raw, 12, out int parsedSecond)) {
            return null;
        }

        int month = raw.Length >= 6 ? parsedMonth : 1;
        int day = raw.Length >= 8 ? parsedDay : 1;
        int hour = raw.Length >= 10 ? parsedHour : 0;
        int minute = raw.Length >= 12 ? parsedMinute : 0;
        int second = raw.Length >= 14 ? parsedSecond : 0;
        TimeSpan offset = TimeSpan.Zero;
        if (raw.Length > 14) {
            if (raw[14] == 'Z') {
                if (raw.Length != 15) return null;
            } else if (raw[14] == '+' || raw[14] == '-') {
                if (!TryParseOffset(raw, out offset)) return null;
                if (raw[14] == '-') offset = -offset;
            } else {
                return null;
            }
        }

        try {
            return new DateTimeOffset(year, month, day, hour, minute, second, offset);
        } catch (ArgumentOutOfRangeException) {
            return null;
        }
    }

    private static bool TryOptionalPart(string value, int startIndex, out int result) {
        result = 0;
        if (value.Length <= startIndex) {
            return true;
        }

        return value.Length >= startIndex + 2 && TryPart(value, startIndex, 2, out result);
    }

    private static bool TryParseOffset(string value, out TimeSpan offset) {
        offset = TimeSpan.Zero;
        if (value.Length != 17 && value.Length != 18 && value.Length != 20 && value.Length != 21) {
            return false;
        }

        if (!TryPart(value, 15, 2, out int hours)) {
            return false;
        }

        int minutes = 0;
        if (value.Length >= 18 && value[17] != '\'') {
            return false;
        }

        if (value.Length >= 20 && !TryPart(value, 18, 2, out minutes)) {
            return false;
        }

        if (hours > 14 || minutes > 59 || (hours == 14 && minutes != 0)) {
            return false;
        }

        if (value.Length == 21 && value[20] != '\'') {
            return false;
        }

        try {
            offset = new TimeSpan(hours, minutes, 0);
            return true;
        } catch (ArgumentOutOfRangeException) {
            return false;
        }
    }

    private static bool TryPart(string value, int startIndex, int length, out int result) {
        result = 0;
        if (startIndex < 0 || length <= 0 || startIndex + length > value.Length) {
            return false;
        }

        for (int index = startIndex; index < startIndex + length; index++) {
            int digit = value[index] - '0';
            if (digit < 0 || digit > 9) {
                result = 0;
                return false;
            }

            result = result * 10 + digit;
        }

        return true;
    }
}
