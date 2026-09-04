using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Parses normalized logical PDF table values without natural-language vocabulary.</summary>
public static class PdfLogicalTableValueParser {
    private static readonly string[] UnambiguousDateTimeFormats = {
        "yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd",
        "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss",
        "yyyy/MM/dd HH:mm", "yyyy/MM/dd HH:mm:ss",
        "yyyy-MM-dd'T'HH:mm", "yyyy-MM-dd'T'HH:mm:ss",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK"
    };

    /// <summary>Parses invariant <c>true</c> or <c>false</c> Boolean literals.</summary>
    public static bool TryParseBoolean(string? value, out bool result) =>
        bool.TryParse(value?.Trim(), out result);

    /// <summary>Parses a number followed by a Unicode percent sign and returns its fractional value.</summary>
    public static bool TryParsePercentage(string? value, CultureInfo? culture, out decimal result) {
        string normalized = value?.Trim() ?? string.Empty;
        if (normalized.Length == 0 || !PdfLogicalTableAnalysis.IsPercentSign(normalized[normalized.Length - 1])) {
            result = 0M;
            return false;
        }
        if (PdfLogicalTableAnalysis.TryParseNumericValue(
                normalized.Substring(0, normalized.Length - 1),
                culture,
                out decimal number)) {
            result = number / 100M;
            return true;
        }
        result = 0M;
        return false;
    }

    /// <summary>Parses a clock time using invariant culture unless an explicit culture is supplied.</summary>
    public static bool TryParseTime(string? value, CultureInfo? culture, out TimeSpan result) {
        string normalized = value?.Trim() ?? string.Empty;
        if (normalized.Length == 0 || normalized.IndexOf(':') < 0) {
            result = default;
            return false;
        }
        if (DateTime.TryParse(
                normalized,
                culture ?? CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault,
                out DateTime parsed) &&
            parsed.Date == DateTime.MinValue.Date) {
            result = parsed.TimeOfDay;
            return true;
        }
        result = default;
        return false;
    }

    /// <summary>
    /// Parses an unambiguous invariant year-first date or date-time. Supplying a culture additionally enables
    /// localized date parsing when the value contains an explicit four-digit year.
    /// </summary>
    public static bool TryParseDateTime(string? value, CultureInfo? culture, out DateTime result) {
        string normalized = value?.Trim() ?? string.Empty;
        if (TryParseUnambiguousDateTime(normalized, out result)) return true;
        return culture is not null &&
            HasExplicitFourDigitYear(normalized) &&
            HasDateComponentBeyondYear(normalized) &&
            DateTime.TryParse(normalized, culture, DateTimeStyles.AllowWhiteSpaces, out result);
    }

    internal static bool LooksLikePlausibleNumericDate(string? value) {
        string normalized = value?.Trim() ?? string.Empty;
        int first = 0;
        int second = 0;
        int third = 0;
        int firstDigits = 0;
        int thirdDigits = 0;
        int componentIndex = 0;
        int index = 0;
        char separator = '\0';
        while (componentIndex < 3) {
            int digits = 0;
            int number = 0;
            while (index < normalized.Length) {
                int digit = CharUnicodeInfo.GetDecimalDigitValue(normalized, index);
                if (digit < 0) break;
                if (digits >= 4) return false;
                number = (number * 10) + digit;
                digits++;
                index += char.IsSurrogatePair(normalized, index) ? 2 : 1;
            }
            if (digits == 0) return false;
            if (componentIndex == 0) {
                first = number;
                firstDigits = digits;
            } else if (componentIndex == 1) {
                second = number;
            } else {
                third = number;
                thirdDigits = digits;
            }
            componentIndex++;
            if (componentIndex == 3) break;
            if (index >= normalized.Length || normalized[index] is not ('.' or '/' or '-')) return false;
            if (separator == '\0') separator = normalized[index];
            else if (normalized[index] != separator) return false;
            index++;
        }
        if (index != normalized.Length) return false;

        bool yearFirst = firstDigits == 4 && first >= 1000 &&
            second is >= 1 and <= 12 && third is >= 1 and <= 31;
        bool yearLast = thirdDigits == 4 && third >= 1000 &&
            ((first is >= 1 and <= 31 && second is >= 1 and <= 12) ||
             (first is >= 1 and <= 12 && second is >= 1 and <= 31));
        return yearFirst || yearLast;
    }

    private static bool TryParseUnambiguousDateTime(string value, out DateTime result) =>
        DateTime.TryParseExact(
            value,
            UnambiguousDateTimeFormats,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces,
            out result);

    private static bool HasExplicitFourDigitYear(string value) {
        int digits = 0;
        int number = 0;
        for (int index = 0; index <= value.Length;) {
            int digit = index < value.Length
                ? CharUnicodeInfo.GetDecimalDigitValue(value, index)
                : -1;
            if (digit >= 0) {
                digits++;
                if (digits <= 4) number = (number * 10) + digit;
                index += char.IsSurrogatePair(value, index) ? 2 : 1;
                continue;
            }
            if (digits == 4 && number >= 1000) return true;
            digits = 0;
            number = 0;
            index++;
        }
        return false;
    }

    private static bool HasDateComponentBeyondYear(string value) {
        int numericComponents = 0;
        bool hasLetter = false;
        bool insideDigits = false;
        for (int index = 0; index < value.Length;) {
            int digit = CharUnicodeInfo.GetDecimalDigitValue(value, index);
            if (digit >= 0) {
                if (!insideDigits) numericComponents++;
                insideDigits = true;
                index += char.IsSurrogatePair(value, index) ? 2 : 1;
                continue;
            }
            insideDigits = false;
            if (char.IsLetter(value, index)) hasLetter = true;
            index += char.IsSurrogatePair(value, index) ? 2 : 1;
        }
        return numericComponents >= 2 || hasLetter;
    }
}
