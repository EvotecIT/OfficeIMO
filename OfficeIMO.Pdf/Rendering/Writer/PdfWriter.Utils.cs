using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static readonly char[] WordSplitChars = new[] { ' ', '\n', '\t' };

    private static string F(double d) {
        if (Math.Abs(d) < 0.0005D) {
            d = 0D;
        }

        return d.ToString("0.###", CultureInfo.InvariantCulture);
    }
    private static string F0(double d) => d.ToString("0", CultureInfo.InvariantCulture);

    private static bool LooksNumeric(string s) {
        if (string.IsNullOrWhiteSpace(s)) return false;
        s = s.Trim();
        if (s.Length >= 2 && s[0] == '(' && s[s.Length - 1] == ')') {
            s = s.Substring(1, s.Length - 2).Trim();
        }

        s = StripCurrencyCode(s);

        int digits = 0;
        foreach (char ch in s) {
            if (char.IsDigit(ch)) digits++;
            else if (ch == ',' ||
                     ch == '.' ||
                     ch == '\'' ||
                     ch == ' ' ||
                     ch == '\u00A0' ||
                     ch == '\u202F' ||
                     ch == '+' ||
                     ch == '-' ||
                     ch == '%' ||
                     IsCurrencySymbol(ch)) continue;
            else return false;
        }
        return digits > 0;
    }

    // Amounts such as "5,147.50 EUR" carry an ISO 4217 code after the number. A leading word is
    // not treated as a code, so labels such as "VAT 23%" keep their own alignment.
    private static string StripCurrencyCode(string s) {
        if (s.Length > 4 && s[s.Length - 4] == ' ' && IsUpperAsciiCode(s, s.Length - 3)) return s.Substring(0, s.Length - 4).TrimEnd();
        return s;
    }

    private static bool IsUpperAsciiCode(string s, int start) {
        for (int i = start; i < start + 3; i++) {
            if (s[i] < 'A' || s[i] > 'Z') return false;
        }
        return true;
    }

    private static bool IsCurrencySymbol(char ch) =>
        ch == '$' ||
        ch == '€' ||
        ch == '£' ||
        ch == '¥' ||
        ch == '¢' ||
        ch == '¤';
}

