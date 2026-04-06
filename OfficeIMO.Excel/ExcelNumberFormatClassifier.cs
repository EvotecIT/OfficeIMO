namespace OfficeIMO.Excel {
    internal static class ExcelNumberFormatClassifier {
        internal static bool LooksLikeDateFormat(string? code) {
            if (string.IsNullOrWhiteSpace(code)) {
                return false;
            }

            string normalized = StripLiteralsAndEscapes(code!);
            if (string.IsNullOrWhiteSpace(normalized)) {
                return false;
            }

            string lower = normalized.ToLowerInvariant();
            if (lower.Contains("am/pm", StringComparison.Ordinal) || lower.Contains("a/p", StringComparison.Ordinal)) {
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

        private static string StripLiteralsAndEscapes(string code) {
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
                        i = ProcessBracketSection(code, i, builder);
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

        private static int ProcessBracketSection(string code, int bracketStart, System.Text.StringBuilder builder) {
            int close = code.IndexOf(']', bracketStart + 1);
            if (close < 0) {
                return code.Length - 1;
            }

            string token = code.Substring(bracketStart + 1, close - bracketStart - 1).Trim().ToLowerInvariant();
            if (token.Length > 0 && token.All(ch => ch is 'h' or 'm' or 's')) {
                builder.Append(token);
            }

            return close;
        }
    }
}
