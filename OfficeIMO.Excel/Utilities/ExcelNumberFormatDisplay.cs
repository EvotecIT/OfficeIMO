using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    internal static class ExcelNumberFormatDisplay {
        private const char LiteralPunctuationMarker = '\u0001';

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

        private static string? FormatNumberValue(double value, uint numberFormatId, string formatCode) {
            int preferredSection = value < 0 ? 1 : value == 0 ? 2 : 0;
            string section = SelectNumberFormatSection(formatCode, preferredSection, value, out int selectedSection);
            if (section.Length == 0) {
                return string.Empty;
            }

            string normalized = StripNumberFormatDecorations(section);
            string lower = normalized.ToLowerInvariant();

            if (numberFormatId == 49U || lower.Contains("@")) {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            if (!ContainsNumericPlaceholder(normalized)) {
                string literal = CleanLiteralAffix(normalized);
                return string.IsNullOrEmpty(literal) ? null : literal;
            }

            if (IsZeroValue(value) && HasOnlyOptionalDigitPlaceholders(normalized)) {
                return string.Empty;
            }

            if (TryFormatFraction(value, normalized, selectedSection, out string fractionText)) {
                return ApplyFractionAffixes(normalized, fractionText);
            }

            if (lower.Contains("e+") || lower.Contains("e-")) {
                int decimals = CountDecimalPlaces(lower);
                double scientificValue = selectedSection == 1 ? Math.Abs(value) : value;
                string scientificText = scientificValue.ToString(BuildScientificFormat(lower, decimals), CultureInfo.InvariantCulture);
                return ApplyNumericAffixes(normalized, scientificText);
            }

            int percentPlaceholders = CountPercentPlaceholders(section);
            bool thousands = lower.Contains("#,##") || lower.Contains(",##");
            bool currency = normalized.IndexOf('$') >= 0
                || normalized.IndexOf('\u20AC') >= 0
                || normalized.IndexOf('\u00A3') >= 0;
            DecimalPlaceInfo decimalPlaces = GetDecimalPlaceInfo(lower);
            double displayValue = value;
            for (int i = 0; i < percentPlaceholders; i++) {
                displayValue *= 100D;
            }

            int scalingCommas = CountScalingCommas(normalized);
            for (int i = 0; i < scalingCommas; i++) {
                displayValue /= 1000D;
            }

            if (selectedSection == 1) {
                displayValue = Math.Abs(displayValue);
            }

            string numericFormat = thousands || currency
                ? "N" + decimalPlaces.Maximum.ToString(CultureInfo.InvariantCulture)
                : "F" + decimalPlaces.Maximum.ToString(CultureInfo.InvariantCulture);
            string text = displayValue.ToString(numericFormat, CultureInfo.InvariantCulture);
            if (decimalPlaces.Optional > 0) {
                text = TrimOptionalDecimalPlaces(text, decimalPlaces.Required);
            }

            return ApplyNumericAffixes(normalized, text);
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

        private static bool TryFormatFraction(double value, string formatCode, int selectedSection, out string text) {
            text = string.Empty;
            int slash = formatCode.IndexOf('/');
            if (slash < 0 || !HasFractionNumerator(formatCode, slash)) {
                return false;
            }

            bool fixedDenominator = TryGetFixedFractionDenominator(formatCode, slash, out int fixedValue);
            int maxDenominator = fixedDenominator
                ? fixedValue
                : GetFractionDenominatorLimit(formatCode, slash);
            if (maxDenominator <= 0) {
                return false;
            }

            double displayValue = selectedSection == 1 ? Math.Abs(value) : value;
            bool negative = displayValue < 0D;
            double absoluteValue = Math.Abs(displayValue);
            if (absoluteValue >= long.MaxValue) {
                return false;
            }

            long whole = (long)Math.Floor(absoluteValue);
            double fractional = absoluteValue - whole;
            long numerator = 0L;
            long denominator = 1L;

            if (fractional > 0.0000000001D) {
                if (fixedDenominator) {
                    denominator = fixedValue;
                    numerator = (int)Math.Round(fractional * denominator, MidpointRounding.AwayFromZero);
                } else {
                    double bestError = double.MaxValue;
                    for (int currentDenominator = 1; currentDenominator <= maxDenominator; currentDenominator++) {
                        long currentNumerator = (long)Math.Round(fractional * currentDenominator, MidpointRounding.AwayFromZero);
                        double candidate = currentNumerator / (double)currentDenominator;
                        double error = Math.Abs(candidate - fractional);
                        if (error + 0.0000000001D < bestError) {
                            bestError = error;
                            numerator = currentNumerator;
                            denominator = currentDenominator;
                        }
                    }
                }

                if (numerator >= denominator) {
                    if (whole > long.MaxValue - (numerator / denominator)) {
                        return false;
                    }

                    whole += numerator / denominator;
                    numerator %= denominator;
                }
            }

            bool mixedFraction = HasMixedFractionWholePart(formatCode, slash);
            if (!mixedFraction && numerator > 0) {
                if (whole > (long.MaxValue - numerator) / denominator) {
                    return false;
                }

                numerator += whole * denominator;
                whole = 0;
            }

            string numericText;
            if (numerator == 0) {
                numericText = whole.ToString(CultureInfo.InvariantCulture);
            } else if (whole > 0) {
                numericText = whole.ToString(CultureInfo.InvariantCulture) + " " +
                    numerator.ToString(CultureInfo.InvariantCulture) + "/" +
                    denominator.ToString(CultureInfo.InvariantCulture);
            } else {
                numericText = numerator.ToString(CultureInfo.InvariantCulture) + "/" +
                    denominator.ToString(CultureInfo.InvariantCulture);
            }

            text = negative && numericText != "0"
                ? "-" + numericText
                : numericText;
            return true;
        }

        private static bool HasFractionNumerator(string formatCode, int slash) {
            for (int i = slash - 1; i >= 0; i--) {
                char ch = formatCode[i];
                if (IsNumericPlaceholder(ch)) {
                    return true;
                }

                if (!char.IsWhiteSpace(ch)) {
                    continue;
                }
            }

            return false;
        }

        private static bool HasMixedFractionWholePart(string formatCode, int slash) {
            int lastSpace = formatCode.LastIndexOf(' ', Math.Max(0, slash - 1), slash);
            if (lastSpace <= 0) {
                return false;
            }

            for (int i = 0; i < lastSpace; i++) {
                if (IsNumericPlaceholder(formatCode[i])) {
                    return true;
                }
            }

            return false;
        }

        private static int GetFractionDenominatorLimit(string formatCode, int slash) {
            int places = 0;
            for (int i = slash + 1; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (IsNumericPlaceholder(ch)) {
                    places++;
                    continue;
                }

                if (places > 0) {
                    break;
                }
            }

            if (places <= 0) {
                return 0;
            }

            int limit = 1;
            for (int i = 0; i < Math.Min(places, 4); i++) {
                limit *= 10;
            }

            return limit - 1;
        }

        private static bool TryGetFixedFractionDenominator(string formatCode, int slash, out int denominator) {
            denominator = 0;
            var builder = new StringBuilder();
            for (int i = slash + 1; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (char.IsDigit(ch)) {
                    builder.Append(ch);
                    continue;
                }

                if (builder.Length > 0) {
                    break;
                }

                if (!char.IsWhiteSpace(ch)) {
                    return false;
                }
            }

            return builder.Length > 0
                && int.TryParse(builder.ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out denominator)
                && denominator > 0;
        }

        private static int CountDecimalPlaces(string formatCode) {
            return GetDecimalPlaceInfo(formatCode).Maximum;
        }

        private static string BuildScientificFormat(string formatCode, int decimals) {
            int exponentIndex = formatCode.IndexOf('e');
            int exponentDigits = 0;
            if (exponentIndex >= 0) {
                for (int i = exponentIndex + 1; i < formatCode.Length; i++) {
                    char ch = formatCode[i];
                    if (ch == '+' || ch == '-') {
                        continue;
                    }

                    if (ch == '0') {
                        exponentDigits++;
                        continue;
                    }

                    break;
                }
            }

            exponentDigits = Math.Max(1, exponentDigits);
            return "0"
                + (decimals > 0 ? "." + new string('0', decimals) : string.Empty)
                + "E+"
                + new string('0', exponentDigits);
        }

        private static DecimalPlaceInfo GetDecimalPlaceInfo(string formatCode) {
            int dot = formatCode.IndexOf('.');
            if (dot < 0) {
                return new DecimalPlaceInfo(0, 0);
            }

            int required = 0;
            int maximum = 0;
            for (int i = dot + 1; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '0') {
                    required++;
                    maximum++;
                    continue;
                }

                if (ch == '#' || ch == '?') {
                    maximum++;
                    continue;
                }

                break;
            }

            return new DecimalPlaceInfo(required, maximum);
        }

        private static string TrimOptionalDecimalPlaces(string text, int requiredDecimalPlaces) {
            int dot = text.IndexOf('.');
            if (dot < 0) {
                return text;
            }

            int end = text.Length - 1;
            while (end > dot + requiredDecimalPlaces && text[end] == '0') {
                end--;
            }

            if (end == dot) {
                return text.Substring(0, dot);
            }

            return end == text.Length - 1
                ? text
                : text.Substring(0, end + 1);
        }

        private static int CountScalingCommas(string formatCode) {
            int last = FindLastNumericPlaceholder(formatCode);
            if (last < 0 || last + 1 >= formatCode.Length) {
                return 0;
            }

            int count = 0;
            for (int i = last + 1; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (char.IsWhiteSpace(ch)) {
                    continue;
                }

                if (ch == ',') {
                    count++;
                    continue;
                }

                break;
            }

            return count;
        }

        private static string SelectNumberFormatSection(string formatCode, int preferredSection) =>
            SelectNumberFormatSection(formatCode, preferredSection, out _);

        private static string SelectNumberFormatSection(string formatCode, int preferredSection, out int selectedSection) =>
            SelectNumberFormatSection(formatCode, preferredSection, value: null, out selectedSection);

        private static string SelectNumberFormatSection(string formatCode, int preferredSection, double? value, out int selectedSection) {
            string[] sections = SplitNumberFormatSections(formatCode);
            selectedSection = 0;
            if (sections.Length == 0) {
                return formatCode;
            }

            if (value.HasValue && sections.Any(section => TryGetSectionCondition(section, out _, out _))) {
                int fallbackSection = -1;
                for (int i = 0; i < sections.Length; i++) {
                    if (TryGetSectionCondition(sections[i], out string? op, out double threshold)) {
                        if (MatchesSectionCondition(value.Value, op!, threshold)) {
                            selectedSection = i;
                            return sections[i];
                        }
                    } else if (fallbackSection < 0) {
                        fallbackSection = i;
                    }
                }

                if (fallbackSection >= 0) {
                    selectedSection = fallbackSection;
                    return sections[fallbackSection];
                }
            }

            if (preferredSection >= 0 && preferredSection < sections.Length) {
                selectedSection = preferredSection;
                return sections[preferredSection];
            }

            return sections[0];
        }

        private static string[] SplitNumberFormatSections(string formatCode) {
            var sections = new List<string>();
            var builder = new StringBuilder(formatCode.Length);
            bool inQuote = false;
            for (int i = 0; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '"') {
                    inQuote = !inQuote;
                    builder.Append(ch);
                    continue;
                }

                if (!inQuote && ch == '\\') {
                    builder.Append(ch);
                    if (i + 1 < formatCode.Length) {
                        builder.Append(formatCode[i + 1]);
                        i++;
                    }

                    continue;
                }

                if (!inQuote && ch == ';') {
                    sections.Add(builder.ToString());
                    builder.Clear();
                    continue;
                }

                builder.Append(ch);
            }

            sections.Add(builder.ToString());
            return sections.ToArray();
        }

        private static bool TryGetSectionCondition(string section, out string? op, out double threshold) {
            op = null;
            threshold = 0D;
            int index = 0;
            while (index < section.Length) {
                while (index < section.Length && char.IsWhiteSpace(section[index])) {
                    index++;
                }

                if (index >= section.Length || section[index] != '[') {
                    return false;
                }

                int close = section.IndexOf(']', index + 1);
                if (close < 0) {
                    return false;
                }

                string token = section.Substring(index + 1, close - index - 1).Trim();
                if (TryParseSectionConditionToken(token, out op, out threshold)) {
                    return true;
                }

                index = close + 1;
            }

            return false;
        }

        private static bool TryParseSectionConditionToken(string token, out string? op, out double threshold) {
            op = null;
            threshold = 0D;
            string[] operators = { ">=", "<=", "<>", ">", "<", "=" };
            foreach (string candidate in operators) {
                if (!token.StartsWith(candidate, StringComparison.Ordinal)) {
                    continue;
                }

                string number = token.Substring(candidate.Length).Trim();
                if (double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out threshold)) {
                    op = candidate;
                    return true;
                }
            }

            return false;
        }

        private static bool MatchesSectionCondition(double value, string op, double threshold) =>
            op switch {
                ">=" => value >= threshold,
                "<=" => value <= threshold,
                "<>" => Math.Abs(value - threshold) > 0.0000000001D,
                ">" => value > threshold,
                "<" => value < threshold,
                "=" => Math.Abs(value - threshold) <= 0.0000000001D,
                _ => false
            };

        private static bool ContainsNumericPlaceholder(string formatCode)
            => formatCode.IndexOf('0') >= 0 || formatCode.IndexOf('#') >= 0 || formatCode.IndexOf('?') >= 0;

        private static bool IsZeroValue(double value) => Math.Abs(value) <= 0.0000000001D;

        private static bool HasOnlyOptionalDigitPlaceholders(string formatCode) {
            bool hasOptionalPlaceholder = false;
            for (int i = 0; i < formatCode.Length; i++) {
                if (formatCode[i] == LiteralPunctuationMarker && i + 1 < formatCode.Length) {
                    i++;
                    continue;
                }

                char ch = formatCode[i];
                if (ch == '0') {
                    return false;
                }

                if (ch == '#' || ch == '?') {
                    hasOptionalPlaceholder = true;
                }
            }

            return hasOptionalPlaceholder;
        }

        private static int CountPercentPlaceholders(string formatCode) {
            int count = 0;
            bool inQuote = false;
            for (int i = 0; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                if (inQuote) {
                    continue;
                }

                if (ch == '\\' || ch == '_' || ch == '*') {
                    if (i + 1 < formatCode.Length) {
                        i++;
                    }

                    continue;
                }

                if (ch == '%') {
                    count++;
                }
            }

            return count;
        }

        private static bool ContainsElapsedToken(string? formatCode, string token) {
            if (string.IsNullOrEmpty(formatCode)) {
                return false;
            }

            return formatCode!.IndexOf("[" + token + "]", StringComparison.OrdinalIgnoreCase) >= 0;
        }

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

        private static string ApplyFractionAffixes(string formatCode, string numericText) {
            int first = FindFirstNumericPlaceholder(formatCode);
            int last = FindLastNumericPlaceholder(formatCode);
            int slash = formatCode.IndexOf('/');
            if (slash >= 0) {
                int denominatorEnd = slash;
                for (int i = slash + 1; i < formatCode.Length; i++) {
                    char ch = formatCode[i];
                    if (char.IsWhiteSpace(ch)) {
                        if (denominatorEnd > slash) {
                            break;
                        }

                        continue;
                    }

                    if (char.IsDigit(ch) || IsNumericPlaceholder(ch)) {
                        denominatorEnd = i;
                        continue;
                    }

                    break;
                }

                last = Math.Max(last, denominatorEnd);
            }

            if (first < 0 || last < first) {
                return CleanLiteralAffix(formatCode);
            }

            string prefix = CleanLiteralAffix(formatCode.Substring(0, first));
            string suffix = CleanLiteralAffix(formatCode.Substring(last + 1));
            return prefix + numericText + suffix;
        }

        private static int FindFirstNumericPlaceholder(string formatCode) {
            for (int i = 0; i < formatCode.Length; i++) {
                if (formatCode[i] == LiteralPunctuationMarker && i + 1 < formatCode.Length) {
                    i++;
                    continue;
                }

                if (IsNumericPlaceholder(formatCode[i])) {
                    return i;
                }
            }

            return -1;
        }

        private static int FindLastNumericPlaceholder(string formatCode) {
            for (int i = formatCode.Length - 1; i >= 0; i--) {
                if (i > 0 && formatCode[i - 1] == LiteralPunctuationMarker) {
                    i--;
                    continue;
                }

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
                if (ch == LiteralPunctuationMarker && i + 1 < value.Length) {
                    builder.Append(value[i + 1]);
                    i++;
                    continue;
                }

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

                if (inQuote && (ch == ',' || ch == '.')) {
                    builder.Append(LiteralPunctuationMarker).Append(ch);
                    continue;
                }

                if (inQuote && IsNumericPlaceholder(ch)) {
                    builder.Append(LiteralPunctuationMarker).Append(ch);
                    continue;
                }

                if (!inQuote && ch == '[') {
                    int close = formatCode.IndexOf(']', i + 1);
                    if (close >= 0) {
                        string token = formatCode.Substring(i + 1, close - i - 1);
                        if (token.All(c => c == 'h' || c == 'H' || c == 'm' || c == 'M' || c == 's' || c == 'S')) {
                            builder.Append('[').Append(token).Append(']');
                        } else if (TryGetBracketedCurrencySymbol(token, out string? symbol)) {
                            builder.Append(symbol);
                        }

                        i = close;
                        continue;
                    }
                }

                if (!inQuote && ch == '\\') {
                    if (i + 1 < formatCode.Length) {
                        char escaped = formatCode[i + 1];
                        if (escaped == ',' || escaped == '.') {
                            builder.Append(LiteralPunctuationMarker);
                        }

                        builder.Append(escaped);
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

        private static bool TryGetBracketedCurrencySymbol(string token, out string? symbol) {
            symbol = null;
            if (token.Length < 2 || token[0] != '$') {
                return false;
            }

            string candidate = token.Substring(1);
            int cultureSeparator = candidate.IndexOf('-');
            if (cultureSeparator >= 0) {
                candidate = candidate.Substring(0, cultureSeparator);
            }

            if (candidate.Length == 0) {
                return false;
            }

            symbol = candidate;
            return true;
        }

        private readonly struct DecimalPlaceInfo {
            internal DecimalPlaceInfo(int required, int maximum) {
                Required = required;
                Maximum = maximum;
            }

            internal int Required { get; }

            internal int Maximum { get; }

            internal int Optional => Maximum - Required;
        }
    }
}
