namespace OfficeIMO.Excel {
    internal static class ExcelNumberFormatClassifier {
        internal static bool LooksLikeDateFormat(string? code) {
            if (string.IsNullOrWhiteSpace(code)) {
                return false;
            }

            string normalized = StripLiteralsAndEscapes(code!, includeElapsedBracketTokens: true);
            if (string.IsNullOrWhiteSpace(normalized)) {
                return false;
            }

            string lower = normalized.ToLowerInvariant();
            if (lower.IndexOf("am/pm", StringComparison.Ordinal) >= 0 || lower.IndexOf("a/p", StringComparison.Ordinal) >= 0) {
                return true;
            }

            bool hasDay = false;
            bool hasYear = false;
            bool hasHour = false;
            bool hasSecond = false;
            bool hasMonth = false;

            foreach (char ch in lower) {
                switch (ch) {
                    case 'd':
                        hasDay = true;
                        break;
                    case 'y':
                        hasYear = true;
                        break;
                    case 'h':
                        hasHour = true;
                        break;
                    case 's':
                        hasSecond = true;
                        break;
                    case 'm':
                        hasMonth = true;
                        break;
                }
            }

            if (hasDay || hasYear || hasHour || hasSecond) {
                return true;
            }

            if (!hasMonth) {
                return false;
            }

            string letterTokens = new string(lower.Where(char.IsLetter).ToArray());
            if (letterTokens.Length > 0 && letterTokens.All(ch => ch == 'm')) {
                return true;
            }

            return lower.Contains(':') || lower.Contains('/') || lower.Contains('-');
        }

        internal static bool LooksLikeDateSystemFormat(string? code) {
            if (string.IsNullOrWhiteSpace(code)) {
                return false;
            }

            string normalized = StripLiteralsAndEscapes(code!, includeElapsedBracketTokens: false);
            if (string.IsNullOrWhiteSpace(normalized)) {
                return false;
            }

            string lower = normalized.ToLowerInvariant();
            bool hasDay = false;
            bool hasYear = false;
            bool hasHour = false;
            bool hasSecond = false;
            bool hasMonth = false;

            foreach (char ch in lower) {
                switch (ch) {
                    case 'd':
                        hasDay = true;
                        break;
                    case 'y':
                        hasYear = true;
                        break;
                    case 'h':
                        hasHour = true;
                        break;
                    case 's':
                        hasSecond = true;
                        break;
                    case 'm':
                        hasMonth = true;
                        break;
                }
            }

            if (hasDay || hasYear) {
                return true;
            }

            if (!hasMonth) {
                return false;
            }

            bool hasMeridiem = lower.IndexOf("am/pm", StringComparison.Ordinal) >= 0 || lower.IndexOf("a/p", StringComparison.Ordinal) >= 0;
            if (!hasHour && !hasSecond && !hasMeridiem) {
                string letterTokens = new string(lower.Where(char.IsLetter).ToArray());
                return letterTokens.Length > 0 && letterTokens.All(ch => ch == 'm');
            }

            return ContainsMonthToken(lower);
        }

        private static string StripLiteralsAndEscapes(string code, bool includeElapsedBracketTokens) {
            var builder = new System.Text.StringBuilder(code.Length);

            for (int i = 0; i < code.Length; i++) {
                char ch = code[i];
                switch (ch) {
                    case '"':
                        i = SkipQuotedLiteral(code, i);
                        break;
                    case '\\':
                    case '_':
                    case '*':
                        if (i + 1 < code.Length) {
                            i++;
                        }
                        break;
                    case '[':
                        i = ProcessBracketSection(code, i, builder, includeElapsedBracketTokens);
                        break;
                    default:
                        builder.Append(ch);
                        break;
                }
            }

            return builder.ToString();
        }

        private static int SkipQuotedLiteral(string code, int quoteStart) {
            int i = quoteStart + 1;
            while (i < code.Length) {
                if (code[i] == '"') {
                    return i;
                }
                i++;
            }

            return code.Length - 1;
        }

        private static int ProcessBracketSection(string code, int bracketStart, System.Text.StringBuilder builder, bool includeElapsedBracketTokens) {
            int close = code.IndexOf(']', bracketStart + 1);
            if (close < 0) {
                return code.Length - 1;
            }

            string token = code.Substring(bracketStart + 1, close - bracketStart - 1).Trim().ToLowerInvariant();
            if (includeElapsedBracketTokens && token.Length > 0 && token.All(ch => ch is 'h' or 'm' or 's')) {
                builder.Append(token);
            }

            return close;
        }

        private static bool ContainsMonthToken(string format) {
            for (int index = 0; index < format.Length; index++) {
                if (format[index] != 'm') {
                    continue;
                }

                int start = index;
                while (index + 1 < format.Length && format[index + 1] == 'm') {
                    index++;
                }

                int end = index;
                if (!IsMinuteToken(format, start, end)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsMinuteToken(string format, int start, int end) {
            return IsAdjacentTimeToken(format, end + 1, searchForward: true)
                || IsAdjacentTimeToken(format, start - 1, searchForward: false);
        }

        private static bool IsAdjacentTimeToken(string format, int index, bool searchForward) {
            bool sawColon = false;
            while (index >= 0 && index < format.Length) {
                char ch = format[index];
                if (ch == ':') {
                    sawColon = true;
                    index += searchForward ? 1 : -1;
                    continue;
                }

                if (char.IsWhiteSpace(ch)) {
                    return false;
                }

                if (!char.IsLetter(ch)) {
                    index += searchForward ? 1 : -1;
                    continue;
                }

                return sawColon && (ch == 'h' || ch == 's');
            }

            return false;
        }
    }
}
