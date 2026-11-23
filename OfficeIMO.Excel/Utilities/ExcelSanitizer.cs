namespace OfficeIMO.Excel.Utilities {
    internal static partial class ExcelSanitizer {
#if NET8_0_OR_GREATER
        private static readonly System.Buffers.SearchValues<char> InvalidXmlControls =
            System.Buffers.SearchValues.Create("\u0000\u0001\u0002\u0003\u0004\u0005\u0006\u0007\u0008\u000B\u000C\u000E\u000F\u0010\u0011\u0012\u0013\u0014\u0015\u0016\u0017\u0018\u0019\u001A\u001B\u001C\u001D\u001E\u001F");
#endif
        // Excel (XML 1.0) allows: Tab (0x9), LF (0xA), CR (0xD), and 0x20..0xD7FF, 0xE000..0xFFFD, 0x10000..0x10FFFF
        // Strip other control characters that trigger "Repaired Records: Cell information".
        public static string SanitizeString(string? input, int maxLength = 32767) {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            var s = input!;
            // Fast path: if nothing suspicious, clamp and return
            bool needsClean = false;
#if NET8_0_OR_GREATER
            int controlIndex = s.AsSpan().IndexOfAnyInRange('\0', '\x1F');
            needsClean = controlIndex >= 0 && s.AsSpan(controlIndex).IndexOfAny(InvalidXmlControls) >= 0;
#else
            for (int i = 0; i < s.Length; i++) {
                char c = s[i];
                if ((c < 0x20 && c != '\t' && c != '\n' && c != '\r')) { needsClean = true; break; }
            }
#endif
            if (!needsClean && s.Length <= maxLength) return s;

            var sb = new StringBuilder(s.Length);
            foreach (var ch in s) {
                if (ch == '\t' || ch == '\n' || ch == '\r' || ch >= 0x20) {
                    sb.Append(ch);
                }
                // else: drop invalid XML control chars
            }
            var cleaned = sb.ToString();
            if (cleaned.Length > maxLength) {
                // Trim hard to Excel's max string length
                cleaned = cleaned.Substring(0, maxLength);
            }
            return cleaned;
        }

        // Defensive formula cleanup: strip leading '=' if provided, remove illegal control chars
        public static string SanitizeFormula(string? formula, bool qualifyArrayFunctions = false) {
            if (string.IsNullOrWhiteSpace(formula)) return string.Empty;
            string f = formula!.Trim();
            if (f.StartsWith("=", StringComparison.Ordinal)) f = f.Substring(1);
            string stored = ExcelFormulaStorage.PrepareForXlsx(SanitizeString(f, maxLength: 8192), qualifyArrayFunctions);
            if (stored.Length > 8192)
                throw new ArgumentException("The formula exceeds the XLSX storage limit after function-prefix expansion.", nameof(formula));
            return stored;
        }
    }
}
