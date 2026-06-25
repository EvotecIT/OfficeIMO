using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    internal static class ExcelNumberFormatDisplay {
        internal static string FormatNumericText(
            double value,
            uint numberFormatId,
            string? formatCode,
            string fallback,
            ExcelDateSystem dateSystem = ExcelDateSystem.NineteenHundred) {
            if (numberFormatId == 0U) {
                return fallback;
            }

            string? resolvedFormatCode = ResolveFormatCode(numberFormatId, formatCode);
            if (string.IsNullOrWhiteSpace(resolvedFormatCode) || string.Equals(resolvedFormatCode, "General", StringComparison.OrdinalIgnoreCase)) {
                return fallback;
            }

            string nonEmptyFormatCode = resolvedFormatCode!;
            if (IsDateNumberFormat(numberFormatId, nonEmptyFormatCode)) {
                return FormatDateValue(value, numberFormatId, nonEmptyFormatCode, dateSystem);
            }

            return FormatNumberValue(value, numberFormatId, nonEmptyFormatCode) ?? fallback;
        }

        internal static bool IsDateNumberFormat(uint numberFormatId, string? formatCode)
            => ExcelBuiltInNumberFormats.IsDate(numberFormatId)
            || ExcelNumberFormatClassifier.LooksLikeDateFormat(formatCode);

        private static string? ResolveFormatCode(uint numberFormatId, string? formatCode) =>
            string.IsNullOrWhiteSpace(formatCode)
                ? ExcelBuiltInNumberFormats.GetCode(numberFormatId)
                : formatCode;

        internal static bool TryGetDateSample(uint numberFormatId, string? formatCode, out string sample) {
            sample = string.Empty;
            if (numberFormatId == 46U || (formatCode?.IndexOf("[h]", StringComparison.OrdinalIgnoreCase) ?? -1) >= 0) {
                return false;
            }

            switch (numberFormatId) {
                case 14:
                    sample = "12/31/9999";
                    return true;
                case 15:
                    sample = "30-Sep-99";
                    return true;
                case 16:
                    sample = "30-Sep";
                    return true;
                case 17:
                    sample = "Sep-99";
                    return true;
                case 18:
                    sample = "12:00 PM";
                    return true;
                case 19:
                    sample = "12:00:00 PM";
                    return true;
                case 20:
                    sample = "23:59";
                    return true;
                case 21:
                    sample = "23:59:59";
                    return true;
                case 22:
                    sample = "12/31/9999 23:59";
                    return true;
                case 45:
                    sample = "59:59";
                    return true;
                case 47:
                    sample = "59:59.0";
                    return true;
            }

            if (string.IsNullOrWhiteSpace(formatCode)) {
                return false;
            }

            sample = new DateTime(2099, 12, 31, 23, 59, 59).ToString(TranslateDateFormat(formatCode!), CultureInfo.InvariantCulture);
            return true;
        }

        private static string FormatDateValue(double value, uint numberFormatId, string formatCode, ExcelDateSystem dateSystem) {
            if (numberFormatId == 46U || formatCode.IndexOf("[h]", StringComparison.OrdinalIgnoreCase) >= 0) {
                TimeSpan duration = TimeSpan.FromDays(value);
                int totalHours = (int)Math.Floor(duration.TotalHours);
                return string.Format(CultureInfo.InvariantCulture, "{0}:{1:00}:{2:00}", totalHours, Math.Abs(duration.Minutes), Math.Abs(duration.Seconds));
            }

            DateTime date;
            try {
                date = ExcelDateSystemConverter.FromSerial(value, dateSystem);
            } catch {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            switch (numberFormatId) {
                case 14: return date.ToString("M/d/yyyy", CultureInfo.InvariantCulture);
                case 15: return date.ToString("d-MMM-yy", CultureInfo.InvariantCulture);
                case 16: return date.ToString("d-MMM", CultureInfo.InvariantCulture);
                case 17: return date.ToString("MMM-yy", CultureInfo.InvariantCulture);
                case 18: return date.ToString("h:mm tt", CultureInfo.InvariantCulture);
                case 19: return date.ToString("h:mm:ss tt", CultureInfo.InvariantCulture);
                case 20: return date.ToString("H:mm", CultureInfo.InvariantCulture);
                case 21: return date.ToString("H:mm:ss", CultureInfo.InvariantCulture);
                case 22: return date.ToString("M/d/yyyy H:mm", CultureInfo.InvariantCulture);
                case 45: return date.ToString("mm:ss", CultureInfo.InvariantCulture);
                case 47: return date.ToString("mm:ss.0", CultureInfo.InvariantCulture);
                default:
                    return date.ToString(TranslateDateFormat(formatCode), CultureInfo.InvariantCulture);
            }
        }

        private static string TranslateDateFormat(string formatCode) {
            string section = SelectNumberFormatSection(formatCode, 0);
            string normalized = StripNumberFormatDecorations(section);
            string lower = normalized.ToLowerInvariant();

            if (lower.Contains("yyyy-mm-dd") && lower.Contains("hh:mm:ss")) return "yyyy-MM-dd HH:mm:ss";
            if (lower.Contains("yyyy-mm-dd") && lower.Contains("hh:mm")) return "yyyy-MM-dd HH:mm";
            if (lower.Contains("yyyy-mm-dd")) return "yyyy-MM-dd";
            if (lower.Contains("dd/mm/yyyy")) return "dd/MM/yyyy";
            if (lower.Contains("mm/dd/yyyy")) return "MM/dd/yyyy";
            if (lower.Contains("m/d/yyyy")) return "M/d/yyyy";
            if (lower.Contains("d-mmm-yy")) return "d-MMM-yy";
            if (lower.Contains("mmm-yy")) return "MMM-yy";
            if (lower.Contains("h:mm:ss") && lower.Contains("am/pm")) return "h:mm:ss tt";
            if (lower.Contains("h:mm") && lower.Contains("am/pm")) return "h:mm tt";
            if (lower.Contains("hh:mm:ss")) return "HH:mm:ss";
            if (lower.Contains("h:mm:ss")) return "H:mm:ss";
            if (lower.Contains("hh:mm")) return "HH:mm";
            if (lower.Contains("h:mm")) return "H:mm";
            return "M/d/yyyy";
        }

        private static string? FormatNumberValue(double value, uint numberFormatId, string formatCode) {
            int preferredSection = value < 0 ? 1 : value == 0 ? 2 : 0;
            string section = SelectNumberFormatSection(formatCode, preferredSection, out int selectedSection);
            string normalized = StripNumberFormatDecorations(section);
            string lower = normalized.ToLowerInvariant();

            if (numberFormatId == 49U || lower.Contains("@")) {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            if (!ContainsNumericPlaceholder(normalized)) {
                string literal = CleanLiteralAffix(normalized);
                return string.IsNullOrEmpty(literal) ? null : literal;
            }

            if (lower.Contains("e+")) {
                int decimals = CountDecimalPlaces(lower);
                double scientificValue = selectedSection == 1 ? Math.Abs(value) : value;
                string scientificText = scientificValue.ToString("E" + decimals.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
                return ApplyNumericAffixes(normalized, scientificText);
            }

            bool percent = lower.Contains("%");
            bool thousands = lower.Contains("#,##") || lower.Contains(",##");
            bool currency = normalized.IndexOf('$') >= 0
                || normalized.IndexOf('\u20AC') >= 0
                || normalized.IndexOf('\u00A3') >= 0;
            int decimalPlaces = CountDecimalPlaces(lower);
            double displayValue = percent ? value * 100.0 : value;
            if (selectedSection == 1) {
                displayValue = Math.Abs(displayValue);
            }

            string numericFormat = thousands || currency
                ? "N" + decimalPlaces.ToString(CultureInfo.InvariantCulture)
                : "F" + decimalPlaces.ToString(CultureInfo.InvariantCulture);
            string text = displayValue.ToString(numericFormat, CultureInfo.InvariantCulture);

            return ApplyNumericAffixes(normalized, text);
        }

        private static int CountDecimalPlaces(string formatCode) {
            int dot = formatCode.IndexOf('.');
            if (dot < 0) {
                return 0;
            }

            int count = 0;
            for (int i = dot + 1; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '0' || ch == '#') {
                    count++;
                    continue;
                }

                break;
            }

            return count;
        }

        private static string SelectNumberFormatSection(string formatCode, int preferredSection) =>
            SelectNumberFormatSection(formatCode, preferredSection, out _);

        private static string SelectNumberFormatSection(string formatCode, int preferredSection, out int selectedSection) {
            string[] sections = formatCode.Split(';');
            selectedSection = 0;
            if (sections.Length == 0) {
                return formatCode;
            }

            if (preferredSection >= 0 && preferredSection < sections.Length && !string.IsNullOrWhiteSpace(sections[preferredSection])) {
                selectedSection = preferredSection;
                return sections[preferredSection];
            }

            return sections[0];
        }

        private static bool ContainsNumericPlaceholder(string formatCode)
            => formatCode.IndexOf('0') >= 0 || formatCode.IndexOf('#') >= 0 || formatCode.IndexOf('?') >= 0;

        private static string ApplyNumericAffixes(string formatCode, string numericText) {
            int first = FindFirstNumericPlaceholder(formatCode);
            int last = FindLastNumericPlaceholder(formatCode);
            if (first < 0 || last < first) {
                return CleanLiteralAffix(formatCode);
            }

            string prefix = CleanLiteralAffix(formatCode.Substring(0, first));
            string suffix = CleanLiteralAffix(formatCode.Substring(last + 1));
            return prefix + numericText + suffix;
        }

        private static int FindFirstNumericPlaceholder(string formatCode) {
            for (int i = 0; i < formatCode.Length; i++) {
                if (IsNumericPlaceholder(formatCode[i])) {
                    return i;
                }
            }

            return -1;
        }

        private static int FindLastNumericPlaceholder(string formatCode) {
            for (int i = formatCode.Length - 1; i >= 0; i--) {
                if (IsNumericPlaceholder(formatCode[i])) {
                    return i;
                }
            }

            return -1;
        }

        private static bool IsNumericPlaceholder(char value) => value == '0' || value == '#' || value == '?';

        private static string CleanLiteralAffix(string value) {
            if (string.IsNullOrEmpty(value)) {
                return string.Empty;
            }

            var builder = new StringBuilder(value.Length);
            for (int i = 0; i < value.Length; i++) {
                char ch = value[i];
                if (ch == ',' || ch == '.') {
                    continue;
                }

                builder.Append(ch);
            }

            return builder.ToString();
        }

        private static string StripNumberFormatDecorations(string formatCode) {
            var builder = new StringBuilder(formatCode.Length);
            bool inQuote = false;

            for (int i = 0; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                if (!inQuote && ch == '[') {
                    int close = formatCode.IndexOf(']', i + 1);
                    if (close >= 0) {
                        string token = formatCode.Substring(i + 1, close - i - 1);
                        if (token.All(c => c == 'h' || c == 'H' || c == 'm' || c == 'M' || c == 's' || c == 'S')) {
                            builder.Append('[').Append(token).Append(']');
                        }

                        i = close;
                        continue;
                    }
                }

                if (!inQuote && ch == '\\') {
                    if (i + 1 < formatCode.Length) {
                        builder.Append(formatCode[i + 1]);
                        i++;
                    }

                    continue;
                }

                if (!inQuote && (ch == '_' || ch == '*')) {
                    if (i + 1 < formatCode.Length) {
                        i++;
                    }
                    continue;
                }

                builder.Append(ch);
            }

            return builder.ToString();
        }
    }
}
