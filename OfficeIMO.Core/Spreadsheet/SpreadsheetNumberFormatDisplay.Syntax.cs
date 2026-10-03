using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Spreadsheet {
    internal static partial class SpreadsheetNumberFormatDisplay {
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

        internal static string SelectNumberFormatSection(string formatCode, int preferredSection) =>
            SelectNumberFormatSection(formatCode, preferredSection, out _);

        internal static string SelectNumberFormatSection(string formatCode, int preferredSection, out int selectedSection) =>
            SelectNumberFormatSection(formatCode, preferredSection, value: null, out selectedSection);

        internal static string SelectNumberFormatSection(string formatCode, int preferredSection, double? value, out int selectedSection) {
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
            var builder = new StringBuilder(Math.Min(formatCode.Length, 256));
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
                    // Excel has at most four sections. Never expand arbitrary trailing
                    // segments from an imported number format into an unbounded list.
                    if (sections.Count == 4) return sections.ToArray();
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

        private static string ApplyNumericAffixes(string formatCode, string numericText) {
            int first = FindFirstNumericPlaceholder(formatCode);
            int last = FindLastNumericPlaceholder(formatCode);
            if (first < 0 || last < first) {
                return CleanLiteralAffix(formatCode);
            }

            string prefix = CleanLiteralAffix(formatCode.Substring(0, first));
            string suffix = CleanLiteralAffix(formatCode.Substring(last + 1));
            if (numericText.Length > 0 && numericText[0] == '-' && prefix.Length > 0
                && prefix[0] is '$' or '\u20AC' or '\u00A3') {
                return "-" + prefix + numericText.Substring(1) + suffix;
            }
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

        internal static string StripNumberFormatDecorations(string formatCode) {
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

                if (inQuote && (IsNumericPlaceholder(ch) || ch is 'E' or 'e')) {
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
                        if (escaped is ',' or '.' or 'E' or 'e' || IsNumericPlaceholder(escaped)) {
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

    }
}
