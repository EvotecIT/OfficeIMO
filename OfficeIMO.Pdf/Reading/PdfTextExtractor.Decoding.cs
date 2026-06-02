using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

public static partial class PdfTextExtractor {
    internal static string UnescapePdfLiteral(string s) {
        var sb = new StringBuilder();
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (c == '\\' && i + 1 < s.Length) {
                char n = s[++i];
                if (n >= '0' && n <= '7') {
                    int value = n - '0';
                    int digits = 1;
                    while (digits < 3 && i + 1 < s.Length && s[i + 1] >= '0' && s[i + 1] <= '7') {
                        value = (value * 8) + (s[++i] - '0');
                        digits++;
                    }
                    sb.Append((char)value);
                    continue;
                }
    
                if (n == '\r') {
                    if (i + 1 < s.Length && s[i + 1] == '\n') {
                        i++;
                    }
                    continue;
                }
    
                if (n == '\n') {
                    continue;
                }
    
                sb.Append(n switch {
                    'n' => '\n',
                    'r' => '\r',
                    't' => '\t',
                    'b' => '\b',
                    'f' => '\f',
                    '(' => '(',
                    ')' => ')',
                    '\\' => '\\',
                    _ => n
                });
            } else sb.Append(c);
        }
        return sb.ToString();
    }
    
    internal static string DecodeHexPdfString(string s) {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
    
        var hex = new StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++) {
            char ch = s[i];
            if (!char.IsWhiteSpace(ch)) hex.Append(ch);
        }
    
        if (hex.Length % 2 == 1) hex.Append('0');
    
        var bytes = new byte[hex.Length / 2];
        for (int i = 0; i < bytes.Length; i++) {
            int hi = HexNibble(hex[i * 2]);
            int lo = HexNibble(hex[i * 2 + 1]);
            bytes[i] = (byte)((hi << 4) | lo);
        }
    
        return PdfWinAnsiEncoding.Decode(bytes);
    
        static int HexNibble(char c) {
            if (c >= '0' && c <= '9') return c - '0';
            if (c >= 'a' && c <= 'f') return 10 + (c - 'a');
            if (c >= 'A' && c <= 'F') return 10 + (c - 'A');
            throw new FormatException($"Invalid hex character '{c}'.");
        }
    }
    
    private static string? ExtractStringValue(string obj, string key) {
        int idx = obj.IndexOf(key, StringComparison.Ordinal);
        if (idx < 0) return null;
        int valueStart = idx + key.Length;
        while (valueStart < obj.Length && char.IsWhiteSpace(obj[valueStart])) {
            valueStart++;
        }
    
        if (valueStart >= obj.Length) {
            return null;
        }
    
        if (obj[valueStart] == '(') {
            int close = FindCloseParen(obj, valueStart);
            if (close < 0) return null;
            string raw = obj.Substring(valueStart + 1, close - valueStart - 1);
            return PdfTextString.DecodeLiteral(raw);
        }
    
        if (obj[valueStart] == '<' && (valueStart + 1 >= obj.Length || obj[valueStart + 1] != '<')) {
            int close = obj.IndexOf('>', valueStart + 1);
            if (close < 0) return null;
            string raw = obj.Substring(valueStart + 1, close - valueStart - 1);
            return DecodeMetadataHexString(raw);
        }
    
        return null;
    }
    
    private static int FindCloseParen(string s, int start) {
        int depth = 0;
        for (int i = start; i < s.Length; i++) {
            char c = s[i];
            if (c == '\\') { i++; continue; }
            if (c == '(') depth++;
            else if (c == ')') { depth--; if (depth == 0) return i; }
        }
        return -1;
    }
    
    private static bool TryReadTextAdvance(string line, out double advanceX, out double advanceY) {
        advanceX = 0;
        advanceY = 0;
    
        if (!line.EndsWith(" Td", StringComparison.Ordinal) && line != "Td") {
            return false;
        }
    
        var parts = line.Split(SpaceSplitChars, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 3 || !string.Equals(parts[parts.Length - 1], "Td", StringComparison.Ordinal)) {
            return false;
        }
    
        return double.TryParse(parts[parts.Length - 3], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out advanceX) &&
               double.TryParse(parts[parts.Length - 2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out advanceY);
    }
    
    private static bool TryReadFontSize(string line, out double fontSize) {
        fontSize = 0;
        var parts = line.Split(SpaceSplitChars, StringSplitOptions.RemoveEmptyEntries);
        return parts.Length >= 3 &&
               string.Equals(parts[parts.Length - 1], "Tf", StringComparison.Ordinal) &&
               double.TryParse(parts[parts.Length - 2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out fontSize);
    }
    
    private static bool TryReadHorizontalScale(string line, out double horizontalScale) {
        horizontalScale = 0;
        var parts = line.Split(SpaceSplitChars, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2 ||
            !string.Equals(parts[parts.Length - 1], "Tz", StringComparison.Ordinal) ||
            !double.TryParse(parts[parts.Length - 2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double percent)) {
            return false;
        }
    
        horizontalScale = percent / 100.0;
        return true;
    }
    
    private static bool TryAppendTextArray(string line, double fontSize, double horizontalScale, Action<string> appendText, Action requestSpace) {
        int operatorIndex = line.LastIndexOf("TJ", StringComparison.Ordinal);
        if (operatorIndex < 0) {
            return false;
        }
    
        int arrayStart = line.IndexOf('[');
        int arrayEnd = line.LastIndexOf(']');
        if (arrayStart < 0 || arrayEnd < arrayStart || arrayEnd > operatorIndex) {
            return false;
        }
    
        string items = line.Substring(arrayStart + 1, arrayEnd - arrayStart - 1);
        int i = 0;
        while (i < items.Length) {
            while (i < items.Length && char.IsWhiteSpace(items[i])) {
                i++;
            }
    
            if (i >= items.Length) {
                break;
            }
    
            char ch = items[i];
            if (ch == '(') {
                appendText(ReadLiteral(items, ref i));
                continue;
            }
    
            if (ch == '<') {
                appendText(ReadHex(items, ref i));
                continue;
            }
    
            if (IsNumberStart(ch)) {
                if (TryReadNumber(items, ref i, out double adjustment) &&
                    ToVisualAdvance(adjustment, fontSize, horizontalScale) > Math.Max(1.5, fontSize * 0.24)) {
                    requestSpace();
                }
                continue;
            }
    
            i++;
        }
    
        return true;
    
        static string ReadLiteral(string s, ref int index) {
            index++; // skip (
            int depth = 1;
            bool escaped = false;
            var raw = new StringBuilder();
            while (index < s.Length && depth > 0) {
                char current = s[index++];
                if (escaped) {
                    raw.Append(current);
                    escaped = false;
                    continue;
                }
    
                if (current == '\\') {
                    raw.Append(current);
                    escaped = true;
                    continue;
                }
    
                if (current == '(') {
                    depth++;
                    raw.Append(current);
                    continue;
                }
    
                if (current == ')') {
                    depth--;
                    if (depth > 0) {
                        raw.Append(current);
                    }
                    continue;
                }
    
                raw.Append(current);
            }
    
            return UnescapePdfLiteral(raw.ToString());
        }
    
        static string ReadHex(string s, ref int index) {
            index++; // skip <
            var raw = new StringBuilder();
            while (index < s.Length && s[index] != '>') {
                raw.Append(s[index]);
                index++;
            }
    
            if (index < s.Length && s[index] == '>') {
                index++;
            }
    
            return DecodeHexPdfString(raw.ToString());
        }
    
        static bool TryReadNumber(string s, ref int index, out double value) {
            int start = index;
            if (s[index] == '+' || s[index] == '-') {
                index++;
            }
    
            while (index < s.Length && (char.IsDigit(s[index]) || s[index] == '.')) {
                index++;
            }
    
    #if NET8_0_OR_GREATER
            return double.TryParse(s.AsSpan(start, index - start), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value);
    #else
            return double.TryParse(s.Substring(start, index - start), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value);
    #endif
        }
    
        static bool IsNumberStart(char ch) => ch == '+' || ch == '-' || ch == '.' || char.IsDigit(ch);
        static double ToVisualAdvance(double tjAdjustment, double currentFontSize, double currentHorizontalScale) => -tjAdjustment / 1000.0 * currentFontSize * currentHorizontalScale;
    }
    
    private static bool TryAppendNextLineShowText(string line, Action<string> appendText, Action requestSpace) {
        Match literalQuote = QuoteLiteralRegex.Match(line);
        if (literalQuote.Success) {
            requestSpace();
            appendText(UnescapePdfLiteral(literalQuote.Groups["txt"].Value));
            return true;
        }
    
        Match hexQuote = QuoteHexRegex.Match(line);
        if (hexQuote.Success) {
            requestSpace();
            appendText(DecodeHexPdfString(hexQuote.Groups["txt"].Value));
            return true;
        }
    
        Match literalDoubleQuote = DoubleQuoteLiteralRegex.Match(line);
        if (literalDoubleQuote.Success) {
            requestSpace();
            appendText(UnescapePdfLiteral(literalDoubleQuote.Groups["txt"].Value));
            return true;
        }
    
        Match hexDoubleQuote = DoubleQuoteHexRegex.Match(line);
        if (hexDoubleQuote.Success) {
            requestSpace();
            appendText(DecodeHexPdfString(hexDoubleQuote.Groups["txt"].Value));
            return true;
        }
    
        return false;
    }
    
    private static string DecodeMetadataHexString(string raw) {
        return PdfTextString.DecodeHex(raw);
    }
}
