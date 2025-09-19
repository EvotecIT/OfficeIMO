using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static readonly char[] WordSplitChars = new[] { ' ', '\n', '\t' };

    private static string F(double d) => d.ToString("0.###", CultureInfo.InvariantCulture);
    private static string F0(double d) => d.ToString("0", CultureInfo.InvariantCulture);

    private static bool LooksNumeric(string s) {
        if (string.IsNullOrWhiteSpace(s)) return false;
        s = s.Trim();
#if NET8_0_OR_GREATER
        if (s.StartsWith('$') || s.EndsWith('%')) return true;
#else
        if (s.StartsWith("$", System.StringComparison.Ordinal) || s.EndsWith("%", System.StringComparison.Ordinal)) return true;
#endif
        int digits = 0;
        foreach (char ch in s) {
            if (char.IsDigit(ch)) digits++;
            else if (ch == ',' || ch == '.' || ch == ' ' || ch == '+' || ch == '-' || ch == '$') continue;
            else return false;
        }
        return digits > 0;
    }
}

