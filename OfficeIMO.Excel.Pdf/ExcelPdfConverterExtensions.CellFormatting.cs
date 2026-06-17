using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static string FormatCellValue(object? value, ExcelCellStyleSnapshot? style, string emptyCellText) {
            if (value == null) {
                return emptyCellText;
            }

            string? formatCode = style?.NumberFormatCode;
            if (!string.IsNullOrWhiteSpace(formatCode)) {
                string? formatted = TryFormatCellValue(value, style!, formatCode!);
                if (formatted != null) {
                    return formatted;
                }
            }

            if (value is IFormattable formattable) {
                return formattable.ToString(null, CultureInfo.InvariantCulture) ?? emptyCellText;
            }

            return value.ToString() ?? emptyCellText;
        }

        private static string? TryFormatCellValue(object value, ExcelCellStyleSnapshot style, string formatCode) {
            string normalized = GetNumberFormatSection(formatCode, 0).Trim();
            if (normalized.Length == 0 ||
                string.Equals(normalized, "General", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "@", StringComparison.Ordinal)) {
                return null;
            }

            if (ContainsElapsedToken(normalized)) {
                double elapsedNumber;
                if (value is DateTime elapsedDate) {
                    elapsedNumber = elapsedDate.ToOADate();
                } else if (!TryGetDouble(value, out elapsedNumber)) {
                    elapsedNumber = double.NaN;
                }

                if (!double.IsNaN(elapsedNumber) &&
                    TryFormatElapsedDuration(elapsedNumber, normalized, out string? elapsedText)) {
                    return elapsedText;
                }
            }

            if (value is DateTime dateValue || style.IsDateLike) {
                DateTime date = value is DateTime directDate
                    ? directDate
                    : TryGetDouble(value, out double oaDate) ? DateTime.FromOADate(oaDate) : default;
                if (date != default) {
                    return date.ToString(ToDotNetDateTimeFormat(normalized), CultureInfo.InvariantCulture);
                }
            }

            if (!TryGetDouble(value, out double number)) {
                return null;
            }

            normalized = GetNumberFormatSection(formatCode, GetNumberFormatSectionIndex(number)).Trim();
            if (normalized.Length == 0 ||
                string.Equals(normalized, "General", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "@", StringComparison.Ordinal)) {
                return null;
            }

            if (normalized.IndexOf('%') >= 0) {
                int decimals = CountDecimalPlaces(normalized);
                bool wrapPercent = ShouldWrapNegativeNumber(normalized, number);
                double percentNumber = wrapPercent ? Math.Abs(number) : number;
                string numeric = (percentNumber * 100D).ToString(decimals > 0 ? "N" + decimals.ToString(CultureInfo.InvariantCulture) : "N0", CultureInfo.InvariantCulture);
                if (wrapPercent) {
                    return "(" + numeric + "%)";
                }

                return numeric + "%";
            }

            bool useGrouping = normalized.IndexOf(',') >= 0;
            int decimalPlaces = CountDecimalPlaces(normalized);
            string numberFormat = (useGrouping ? "N" : "F") + decimalPlaces.ToString(CultureInfo.InvariantCulture);
            bool wrapNumber = ShouldWrapNegativeNumber(normalized, number);
            double displayNumber = wrapNumber ? Math.Abs(number) : number;
            string numericValue = ApplyQuotedLiterals(normalized, displayNumber.ToString(numberFormat, CultureInfo.InvariantCulture));
            return wrapNumber ? "(" + numericValue + ")" : numericValue;
        }

        private static bool TryFormatElapsedDuration(double value, string formatCode, out string? text) {
            text = null;
            if (!ContainsElapsedToken(formatCode)) {
                return false;
            }

            bool negative = value < 0D;
            TimeSpan duration = TimeSpan.FromDays(Math.Abs(value));
            string result = formatCode;
            bool replaced = false;
            if (TryReplaceElapsedToken(ref result, "[hh]", (long)Math.Floor(duration.TotalHours), 2) ||
                TryReplaceElapsedToken(ref result, "[h]", (long)Math.Floor(duration.TotalHours), 1)) {
                replaced = true;
            } else if (TryReplaceElapsedToken(ref result, "[mm]", (long)Math.Floor(duration.TotalMinutes), 2) ||
                       TryReplaceElapsedToken(ref result, "[m]", (long)Math.Floor(duration.TotalMinutes), 1)) {
                replaced = true;
            } else if (TryReplaceElapsedToken(ref result, "[ss]", (long)Math.Floor(duration.TotalSeconds), 2) ||
                       TryReplaceElapsedToken(ref result, "[s]", (long)Math.Floor(duration.TotalSeconds), 1)) {
                replaced = true;
            }

            if (!replaced) {
                return false;
            }

            result = ReplaceUnquotedIgnoreCase(result, "hh", duration.Hours.ToString("D2", CultureInfo.InvariantCulture));
            result = ReplaceUnquotedIgnoreCase(result, "h", duration.Hours.ToString(CultureInfo.InvariantCulture));
            result = ReplaceUnquotedIgnoreCase(result, "mm", duration.Minutes.ToString("D2", CultureInfo.InvariantCulture));
            result = ReplaceUnquotedIgnoreCase(result, "m", duration.Minutes.ToString(CultureInfo.InvariantCulture));
            result = ReplaceUnquotedIgnoreCase(result, "ss", duration.Seconds.ToString("D2", CultureInfo.InvariantCulture));
            result = ReplaceUnquotedIgnoreCase(result, "s", duration.Seconds.ToString(CultureInfo.InvariantCulture));
            result = RemoveExcelFormatQuotes(result);
            text = negative ? "-" + result : result;
            return true;
        }

        private static bool ContainsElapsedToken(string formatCode) =>
            formatCode.IndexOf("[h]", StringComparison.OrdinalIgnoreCase) >= 0 ||
            formatCode.IndexOf("[hh]", StringComparison.OrdinalIgnoreCase) >= 0 ||
            formatCode.IndexOf("[m]", StringComparison.OrdinalIgnoreCase) >= 0 ||
            formatCode.IndexOf("[mm]", StringComparison.OrdinalIgnoreCase) >= 0 ||
            formatCode.IndexOf("[s]", StringComparison.OrdinalIgnoreCase) >= 0 ||
            formatCode.IndexOf("[ss]", StringComparison.OrdinalIgnoreCase) >= 0;

        private static bool TryReplaceElapsedToken(ref string formatCode, string token, long value, int minimumDigits) {
            int index = formatCode.IndexOf(token, StringComparison.OrdinalIgnoreCase);
            if (index < 0) {
                return false;
            }

            string replacement = minimumDigits > 1
                ? value.ToString("D" + minimumDigits.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
                : value.ToString(CultureInfo.InvariantCulture);
            formatCode = formatCode.Substring(0, index) + replacement + formatCode.Substring(index + token.Length);
            return true;
        }

        private static bool TryGetDouble(object value, out double number) {
            switch (value) {
                case double doubleValue:
                    number = doubleValue;
                    return true;
                case float floatValue:
                    number = floatValue;
                    return true;
                case decimal decimalValue:
                    number = (double)decimalValue;
                    return true;
                case int intValue:
                    number = intValue;
                    return true;
                case long longValue:
                    number = longValue;
                    return true;
                case short shortValue:
                    number = shortValue;
                    return true;
                case byte byteValue:
                    number = byteValue;
                    return true;
                default:
                    return double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number);
            }
        }

        private static string GetNumberFormatSection(string formatCode, int sectionIndex) {
            if (sectionIndex < 0) {
                sectionIndex = 0;
            }

            int currentSection = 0;
            int sectionStart = 0;
            bool inQuote = false;
            for (int i = 0; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                if (ch == '\\') {
                    i++;
                    continue;
                }

                if (ch != ';' || inQuote) {
                    continue;
                }

                if (currentSection == sectionIndex) {
                    return formatCode.Substring(sectionStart, i - sectionStart);
                }

                currentSection++;
                sectionStart = i + 1;
            }

            if (currentSection == sectionIndex) {
                return formatCode.Substring(sectionStart);
            }

            inQuote = false;
            for (int i = 0; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                if (ch == '\\') {
                    i++;
                    continue;
                }

                if (ch == ';' && !inQuote) {
                    return formatCode.Substring(0, i);
                }
            }

            return formatCode;
        }

        private static int GetNumberFormatSectionIndex(double number) {
            if (number < 0D) {
                return 1;
            }

            return number == 0D ? 2 : 0;
        }

        private static bool ShouldWrapNegativeNumber(string formatCode, double value) =>
            value < 0D && formatCode.IndexOf('(') >= 0 && formatCode.IndexOf(')') > formatCode.IndexOf('(');

        private static int CountDecimalPlaces(string formatCode) {
            int decimalIndex = formatCode.IndexOf('.');
            if (decimalIndex < 0) {
                return 0;
            }

            int count = 0;
            for (int i = decimalIndex + 1; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '0' || ch == '#') {
                    count++;
                    continue;
                }

                break;
            }

            return count;
        }

        private static string ApplyQuotedLiterals(string formatCode, string numericValue) {
            string prefix = string.Empty;
            string suffix = string.Empty;
            int index = 0;
            while (index < formatCode.Length) {
                int quoteStart = formatCode.IndexOf('"', index);
                if (quoteStart < 0) {
                    break;
                }

                int quoteEnd = formatCode.IndexOf('"', quoteStart + 1);
                if (quoteEnd <= quoteStart + 1) {
                    break;
                }

                string literal = formatCode.Substring(quoteStart + 1, quoteEnd - quoteStart - 1);
                bool hasPlaceholderBefore = HasNumberPlaceholder(formatCode, 0, quoteStart);
                if (hasPlaceholderBefore) {
                    if (quoteStart > 0 && char.IsWhiteSpace(formatCode[quoteStart - 1])) {
                        suffix += " ";
                    }

                    suffix += literal;
                } else {
                    prefix += literal;
                    if (quoteEnd + 1 < formatCode.Length && char.IsWhiteSpace(formatCode[quoteEnd + 1])) {
                        prefix += " ";
                    }
                }

                index = quoteEnd + 1;
            }

            return prefix + numericValue + suffix;
        }

        private static string ReplaceUnquotedIgnoreCase(string value, string oldValue, string newValue) {
            var builder = new System.Text.StringBuilder(value.Length);
            bool inQuote = false;
            int index = 0;
            while (index < value.Length) {
                char ch = value[index];
                if (ch == '"') {
                    inQuote = !inQuote;
                    builder.Append(ch);
                    index++;
                    continue;
                }

                if (!inQuote &&
                    index + oldValue.Length <= value.Length &&
                    string.Compare(value, index, oldValue, 0, oldValue.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                    builder.Append(newValue);
                    index += oldValue.Length;
                    continue;
                }

                builder.Append(ch);
                index++;
            }

            return builder.ToString();
        }

        private static string RemoveExcelFormatQuotes(string value) {
            var builder = new System.Text.StringBuilder(value.Length);
            bool inQuote = false;
            for (int i = 0; i < value.Length; i++) {
                char ch = value[i];
                if (ch == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                builder.Append(ch);
            }

            return builder.ToString();
        }

        private static bool HasNumberPlaceholder(string formatCode, int start, int end) {
            for (int i = start; i < end && i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '0' || ch == '#' || ch == '?') {
                    return true;
                }
            }

            return false;
        }

        private static string ToDotNetDateTimeFormat(string excelFormat) {
            string format = StripExcelBracketAndColorTokens(excelFormat);
            format = ReplaceExcelDateTokens(format);
            format = ReplaceIgnoreCase(format, "AM/PM", "tt");
            format = ReplaceIgnoreCase(format, "A/P", "tt");
            return format;
        }

        private static string ReplaceIgnoreCase(string value, string oldValue, string newValue) {
            return value
                .Replace(oldValue, newValue)
                .Replace(oldValue.ToLowerInvariant(), newValue)
                .Replace(oldValue.ToUpperInvariant(), newValue);
        }

        private static string StripExcelBracketAndColorTokens(string format) {
            var builder = new System.Text.StringBuilder(format.Length);
            for (int i = 0; i < format.Length; i++) {
                char ch = format[i];
                if (ch == '[') {
                    int close = format.IndexOf(']', i + 1);
                    if (close >= 0) {
                        string token = format.Substring(i + 1, close - i - 1);
                        if (token.All(c => c == 'h' || c == 'H' || c == 'm' || c == 'M' || c == 's' || c == 'S')) {
                            builder.Append(token);
                        }

                        i = close;
                        continue;
                    }
                }

                if (ch == '\\' || ch == '_') {
                    if (i + 1 < format.Length) {
                        builder.Append(format[i + 1]);
                        i++;
                    }
                    continue;
                }

                builder.Append(ch);
            }

            return builder.ToString();
        }

        private static string ReplaceExcelDateTokens(string format) {
            var builder = new System.Text.StringBuilder(format.Length);
            for (int i = 0; i < format.Length;) {
                char ch = format[i];
                if (ch == '"') {
                    int end = format.IndexOf('"', i + 1);
                    if (end < 0) {
                        break;
                    }

                    builder.Append('\'').Append(format.Substring(i + 1, end - i - 1).Replace("'", "\\'")).Append('\'');
                    i = end + 1;
                    continue;
                }

                if (!IsExcelDateFormatLetter(ch)) {
                    builder.Append(ch);
                    i++;
                    continue;
                }

                int start = i;
                while (i < format.Length && char.ToLowerInvariant(format[i]) == char.ToLowerInvariant(ch)) {
                    i++;
                }

                string token = format.Substring(start, i - start);
                builder.Append(ConvertExcelDateToken(token, builder, format, i));
            }

            return builder.ToString();
        }

        private static bool IsExcelDateFormatLetter(char ch) {
            switch (char.ToLowerInvariant(ch)) {
                case 'y':
                case 'm':
                case 'd':
                case 'h':
                case 's':
                    return true;
                default:
                    return false;
            }
        }

        private static string ConvertExcelDateToken(string token, System.Text.StringBuilder output, string format, int nextIndex) {
            char lower = char.ToLowerInvariant(token[0]);
            switch (lower) {
                case 'y':
                    return token.Length <= 2 ? "yy" : "yyyy";
                case 'd':
                    return token.Length <= 1 ? "d" : token.Length == 2 ? "dd" : token.Length == 3 ? "ddd" : "dddd";
                case 'h':
                    return token.Length <= 1 ? "h" : "hh";
                case 's':
                    return token.Length <= 1 ? "s" : "ss";
                case 'm':
                    bool timeMinute = PreviousNonSpace(output) == ':' || NextNonSpace(format, nextIndex) == ':';
                    if (timeMinute) {
                        return token.Length <= 1 ? "m" : "mm";
                    }

                    return token.Length <= 1 ? "M" : token.Length == 2 ? "MM" : token.Length == 3 ? "MMM" : "MMMM";
                default:
                    return token;
            }
        }

        private static char PreviousNonSpace(System.Text.StringBuilder builder) {
            for (int i = builder.Length - 1; i >= 0; i--) {
                if (!char.IsWhiteSpace(builder[i])) {
                    return builder[i];
                }
            }

            return '\0';
        }

        private static char NextNonSpace(string value, int startIndex) {
            for (int i = startIndex; i < value.Length; i++) {
                if (!char.IsWhiteSpace(value[i])) {
                    return value[i];
                }
            }

            return '\0';
        }

    }
}
