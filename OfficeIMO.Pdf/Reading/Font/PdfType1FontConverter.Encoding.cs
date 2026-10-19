using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

internal static partial class PdfType1FontConverter {
    private static string?[] ReadEncoding(string text) {
        var encoding = new string?[256];
        int cursor = 0;
        while (TryReadEncodingToken(text, ref cursor, out string token)) {
            if (token != "/Encoding") continue;
            if (!TryReadEncodingToken(text, ref cursor, out token)) throw new InvalidDataException("Missing Type 1 encoding.");
            if (token.Length > 0 && char.IsLetter(token[0])) {
                string?[] named = PdfFontEncodingNames.Get(token) ?? throw new InvalidDataException("Unsupported named Type 1 encoding.");
                Array.Copy(named, encoding, 256);
            } else if (token != "256") {
                throw new InvalidDataException("Unsupported Type 1 encoding construction.");
            }
            int depth = 0;
            while (TryReadEncodingToken(text, ref cursor, out token)) {
                if (token is "{" or "[") { depth++; continue; }
                if (token is "}" or "]") { depth--; continue; }
                if (depth != 0) continue;
                if (token == "def") return encoding;
                if (token != "dup") continue;
                // Only executable top-level edits affect the supported encoding vector.
                // Comments, procedure bodies and string literals cannot supply assignments.
                if (!TryReadEncodingToken(text, ref cursor, out string number) ||
                    !int.TryParse(number, NumberStyles.None, CultureInfo.InvariantCulture, out int code) ||
                    !TryReadEncodingToken(text, ref cursor, out string name) || (name.Length == 0 || name[0] != '/') ||
                    !TryReadEncodingToken(text, ref cursor, out string put) || put != "put" || code < 0 || code >= 256)
                    throw new InvalidDataException("Unsupported Type 1 encoding edit.");
                encoding[code] = name.Substring(1);
            }
            throw new InvalidDataException("Unterminated Type 1 encoding.");
        }
        return encoding;
    }

    // Reading cleartext encoding metadata never executes PostScript or converts outlines.
    internal static string?[]? TryReadBuiltinEncoding(byte[] program) {
        try {
            byte[] data = StripPfbHeaders(program);
            int end = IndexOf(data, "eexec", 0);
            return end < 0 ? null : ReadEncoding(Encoding.ASCII.GetString(data, 0, end));
        } catch (InvalidDataException) { return null; }
    }

    private static bool TryReadEncodingToken(string text, ref int cursor, out string token) {
        token = string.Empty;
        while (cursor < text.Length) {
            char current = text[cursor];
            if (char.IsWhiteSpace(current)) { cursor++; continue; }
            if (current == '%') {
                while (cursor < text.Length && text[cursor] is not ('\r' or '\n')) cursor++;
                continue;
            }
            if (current == '(') {
                int depth = 1; cursor++;
                while (cursor < text.Length && depth > 0) {
                    current = text[cursor++];
                    if (current == '\\' && cursor < text.Length) cursor++;
                    else if (current == '(') depth++;
                    else if (current == ')') depth--;
                }
                // A literal is a token, but never an executable name or edit.
                token = "(literal)"; return true;
            }
            if (current == '<') {
                cursor++;
                while (cursor < text.Length && text[cursor] != '>') cursor++;
                if (cursor < text.Length) cursor++;
                token = "<literal>"; return true;
            }
            int start = cursor++;
            if (current is not ('{' or '}' or '[' or ']')) {
                while (cursor < text.Length && !char.IsWhiteSpace(text[cursor]) &&
                    text[cursor] is not ('{' or '}' or '[' or ']' or '(' or ')' or '<' or '>' or '%' or '/')) cursor++;
            }
            token = text.Substring(start, cursor - start);
            return true;
        }
        return false;
    }
}
