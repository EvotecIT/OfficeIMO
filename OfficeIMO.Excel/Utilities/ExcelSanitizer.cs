using System;
using System.Text;

namespace OfficeIMO.Excel.Utilities
{
    internal static class ExcelSanitizer
    {
        // Excel (XML 1.0) allows: Tab (0x9), LF (0xA), CR (0xD), and 0x20..0xD7FF, 0xE000..0xFFFD, 0x10000..0x10FFFF
        // Strip other control characters that trigger "Repaired Records: Cell information".
        public static string SanitizeString(string? input, int maxLength = 32767)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            var s = input!;
            // Fast path: if nothing suspicious, clamp and return
            bool needsClean = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if ((c < 0x20 && c != '\t' && c != '\n' && c != '\r')) { needsClean = true; break; }
            }
            if (!needsClean && s.Length <= maxLength) return s;

            var sb = new StringBuilder(s.Length);
            foreach (var ch in s)
            {
                if (ch == '\t' || ch == '\n' || ch == '\r' || ch >= 0x20)
                {
                    sb.Append(ch);
                }
                // else: drop invalid XML control chars
            }
            var cleaned = sb.ToString();
            if (cleaned.Length > maxLength)
            {
                // Trim hard to Excel's max string length
                cleaned = cleaned.Substring(0, maxLength);
            }
            return cleaned;
        }

        // Defensive formula cleanup: strip leading '=' if provided, remove illegal control chars
        public static string SanitizeFormula(string? formula)
        {
            if (string.IsNullOrWhiteSpace(formula)) return string.Empty;
            string f = formula!.Trim();
            if (f.StartsWith("=", StringComparison.Ordinal)) f = f.Substring(1);
            return SanitizeString(f, maxLength: 8192); // practical limit for formulas
        }
    }
}
