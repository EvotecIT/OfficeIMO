using System.Globalization;
using System.Text;
using OfficeIMO.Spreadsheet;
using static OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay;

namespace OfficeIMO.Excel {
    internal static partial class ExcelNumberFormatDisplay {

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

            return numberFormatId == 49U ? value.ToString(CultureInfo.InvariantCulture)
                : FormatNumericValue(value, nonEmptyFormatCode) ?? fallback;
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
            if (IsElapsedDurationFormat(numberFormatId, formatCode)) {
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

            try {
                sample = new DateTime(2099, 12, 31, 23, 59, 59).ToString(TranslateDateFormat(formatCode!), CultureInfo.InvariantCulture);
                return true;
            } catch (FormatException) {
                return false;
            }
        }

        private static string FormatDateValue(double value, uint numberFormatId, string formatCode, ExcelDateSystem dateSystem) {
            if (TryFormatElapsedDuration(value, numberFormatId, formatCode, out string durationText)) {
                return durationText;
            }

            DateTime date;
            try {
                date = ExcelDateSystemConverter.FromSerial(value, dateSystem);
            } catch {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            try {
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
            } catch (FormatException) {
                return value.ToString(CultureInfo.InvariantCulture);
            }
        }

        private static string TranslateDateFormat(string formatCode) {
            string section = SelectNumberFormatSection(formatCode, 0);
            string normalized = StripNumberFormatDecorations(section);
            string lower = normalized.ToLowerInvariant();

            if (lower.Contains("yyyy-mm-dd") && lower.Contains("hh:mm:ss")) return "yyyy-MM-dd HH:mm:ss";
            if (lower.Contains("yyyy-mm-dd") && lower.Contains("hh:mm")) return "yyyy-MM-dd HH:mm";
            if (lower.Contains("yyyy-mm-dd")) return "yyyy-MM-dd";
            if (lower.Contains("dd/mm/yyyy") && lower.Contains("h:mm:ss") && lower.Contains("am/pm")) return "dd/MM/yyyy h:mm:ss tt";
            if (lower.Contains("dd/mm/yyyy") && lower.Contains("h:mm") && lower.Contains("am/pm")) return "dd/MM/yyyy h:mm tt";
            if (lower.Contains("dd/mm/yyyy") && lower.Contains("h:mm:ss")) return "dd/MM/yyyy H:mm:ss";
            if (lower.Contains("dd/mm/yyyy") && lower.Contains("h:mm")) return "dd/MM/yyyy H:mm";
            if (lower.Contains("dd/mm/yyyy")) return "dd/MM/yyyy";
            if (lower.Contains("mm/dd/yyyy") && lower.Contains("h:mm:ss") && lower.Contains("am/pm")) return "MM/dd/yyyy h:mm:ss tt";
            if (lower.Contains("mm/dd/yyyy") && lower.Contains("h:mm") && lower.Contains("am/pm")) return "MM/dd/yyyy h:mm tt";
            if (lower.Contains("mm/dd/yyyy") && lower.Contains("h:mm:ss")) return "MM/dd/yyyy H:mm:ss";
            if (lower.Contains("mm/dd/yyyy") && lower.Contains("h:mm")) return "MM/dd/yyyy H:mm";
            if (lower.Contains("mm/dd/yyyy")) return "MM/dd/yyyy";
            if (lower.Contains("dd/mm/yy") && lower.Contains("h:mm:ss") && lower.Contains("am/pm")) return "dd/MM/yy h:mm:ss tt";
            if (lower.Contains("dd/mm/yy") && lower.Contains("h:mm") && lower.Contains("am/pm")) return "dd/MM/yy h:mm tt";
            if (lower.Contains("dd/mm/yy") && lower.Contains("h:mm:ss")) return "dd/MM/yy H:mm:ss";
            if (lower.Contains("dd/mm/yy") && lower.Contains("h:mm")) return "dd/MM/yy H:mm";
            if (lower.Contains("dd/mm/yy")) return "dd/MM/yy";
            if (lower.Contains("mm/dd/yy") && lower.Contains("h:mm:ss") && lower.Contains("am/pm")) return "MM/dd/yy h:mm:ss tt";
            if (lower.Contains("mm/dd/yy") && lower.Contains("h:mm") && lower.Contains("am/pm")) return "MM/dd/yy h:mm tt";
            if (lower.Contains("mm/dd/yy") && lower.Contains("h:mm:ss")) return "MM/dd/yy H:mm:ss";
            if (lower.Contains("mm/dd/yy") && lower.Contains("h:mm")) return "MM/dd/yy H:mm";
            if (lower.Contains("mm/dd/yy")) return "MM/dd/yy";
            if (lower.Contains("m/d/yyyy") && lower.Contains("h:mm:ss") && lower.Contains("am/pm")) return "M/d/yyyy h:mm:ss tt";
            if (lower.Contains("m/d/yyyy") && lower.Contains("h:mm") && lower.Contains("am/pm")) return "M/d/yyyy h:mm tt";
            if (lower.Contains("m/d/yyyy") && lower.Contains("h:mm:ss")) return "M/d/yyyy H:mm:ss";
            if (lower.Contains("m/d/yyyy") && lower.Contains("h:mm")) return "M/d/yyyy H:mm";
            if (lower.Contains("m/d/yyyy")) return "M/d/yyyy";
            if (lower.Contains("m/d/yy") && lower.Contains("h:mm:ss") && lower.Contains("am/pm")) return "M/d/yy h:mm:ss tt";
            if (lower.Contains("m/d/yy") && lower.Contains("h:mm") && lower.Contains("am/pm")) return "M/d/yy h:mm tt";
            if (lower.Contains("m/d/yy") && lower.Contains("h:mm:ss")) return "M/d/yy H:mm:ss";
            if (lower.Contains("m/d/yy") && lower.Contains("h:mm")) return "M/d/yy H:mm";
            if (lower.Contains("m/d/yy")) return "M/d/yy";
            if (HasNamedDateToken(lower)) return TranslateNamedDateFormat(normalized);
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

        private static bool HasNamedDateToken(string formatCode) =>
            formatCode.IndexOf("mmm", StringComparison.OrdinalIgnoreCase) >= 0
            || formatCode.IndexOf("dddd", StringComparison.OrdinalIgnoreCase) >= 0
            || formatCode.IndexOf("ddd", StringComparison.OrdinalIgnoreCase) >= 0;

        private static string TranslateNamedDateFormat(string formatCode) {
            string lower = formatCode.ToLowerInvariant();
            bool twelveHour = lower.Contains("am/pm");
            var builder = new StringBuilder(formatCode.Length);
            for (int i = 0; i < formatCode.Length;) {
                if (i + 5 <= formatCode.Length && string.Equals(formatCode.Substring(i, 5), "am/pm", StringComparison.OrdinalIgnoreCase)) {
                    builder.Append("tt");
                    i += 5;
                    continue;
                }

                char ch = formatCode[i];
                char token = char.ToLowerInvariant(ch);
                if (token is 'd' or 'm' or 'y' or 'h' or 's') {
                    int start = i;
                    while (i < formatCode.Length && char.ToLowerInvariant(formatCode[i]) == token) {
                        i++;
                    }

                    int length = i - start;
                    builder.Append(TranslateDateToken(formatCode, start, length, token, twelveHour));
                    continue;
                }

                if (ch != LiteralPunctuationMarker) {
                    builder.Append(ch);
                }

                i++;
            }

            return builder.ToString();
        }

        private static string TranslateDateToken(string formatCode, int start, int length, char token, bool twelveHour) =>
            token switch {
                'd' => length >= 4 ? "dddd" : length == 3 ? "ddd" : length == 2 ? "dd" : "d",
                'm' when length >= 4 => "MMMM",
                'm' when length == 3 => "MMM",
                'm' when IsMinuteToken(formatCode, start, length) => length == 1 ? "m" : "mm",
                'm' => length == 1 ? "M" : "MM",
                'y' => length >= 4 ? "yyyy" : "yy",
                'h' => twelveHour ? length == 1 ? "h" : "hh" : length == 1 ? "H" : "HH",
                's' => length == 1 ? "s" : "ss",
                _ => new string(token, length)
            };

        private static bool IsMinuteToken(string formatCode, int start, int length) {
            int before = start - 1;
            while (before >= 0 && char.IsWhiteSpace(formatCode[before])) {
                before--;
            }

            int after = start + length;
            while (after < formatCode.Length && char.IsWhiteSpace(formatCode[after])) {
                after++;
            }

            return (before >= 0 && formatCode[before] == ':')
                || (after < formatCode.Length && formatCode[after] == ':');
        }

        private static bool IsElapsedDurationFormat(uint numberFormatId, string? formatCode)
            => numberFormatId == 46U
            || ContainsElapsedToken(formatCode, "h")
            || ContainsElapsedToken(formatCode, "hh")
            || ContainsElapsedToken(formatCode, "m")
            || ContainsElapsedToken(formatCode, "mm")
            || ContainsElapsedToken(formatCode, "s")
            || ContainsElapsedToken(formatCode, "ss");

        private static bool TryFormatElapsedDuration(double value, uint numberFormatId, string formatCode, out string text) {
            text = string.Empty;
            string section = SelectNumberFormatSection(formatCode, 0);
            string normalized = StripNumberFormatDecorations(section);
            string lower = normalized.ToLowerInvariant();

            bool hasHours = numberFormatId == 46U || ContainsElapsedToken(lower, "h") || ContainsElapsedToken(lower, "hh");
            bool hasMinutes = ContainsElapsedToken(lower, "m") || ContainsElapsedToken(lower, "mm");
            bool hasSeconds = ContainsElapsedToken(lower, "s") || ContainsElapsedToken(lower, "ss");
            if (!hasHours && !hasMinutes && !hasSeconds) {
                return false;
            }

            TimeSpan duration;
            TimeSpan absolute;
            try {
                duration = TimeSpan.FromDays(value);
                absolute = duration.Duration();
            } catch (ArgumentException) {
                return false;
            } catch (OverflowException) {
                return false;
            }

            bool negative = duration.Ticks < 0;
            string sign = negative ? "-" : string.Empty;

            if (hasHours) {
                long totalHours = (long)Math.Floor(absolute.TotalHours);
                if (lower.Contains(":mm:ss")) {
                    text = string.Format(CultureInfo.InvariantCulture, "{0}{1}:{2:00}:{3:00}", sign, totalHours, absolute.Minutes, absolute.Seconds);
                } else if (lower.Contains(":mm")) {
                    text = string.Format(CultureInfo.InvariantCulture, "{0}{1}:{2:00}", sign, totalHours, absolute.Minutes);
                } else {
                    text = sign + totalHours.ToString(CultureInfo.InvariantCulture);
                }

                return true;
            }

            if (hasMinutes) {
                long totalMinutes = (long)Math.Floor(absolute.TotalMinutes);
                if (lower.Contains(":ss")) {
                    text = string.Format(CultureInfo.InvariantCulture, "{0}{1}:{2:00}", sign, totalMinutes, absolute.Seconds);
                } else {
                    text = sign + totalMinutes.ToString(CultureInfo.InvariantCulture);
                }

                return true;
            }

            long totalSeconds = (long)Math.Floor(absolute.TotalSeconds);
            text = sign + totalSeconds.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        private static bool ContainsElapsedToken(string? formatCode, string token) {
            if (string.IsNullOrEmpty(formatCode)) {
                return false;
            }

            return formatCode!.IndexOf("[" + token + "]", StringComparison.OrdinalIgnoreCase) >= 0;
        }

    }
}
