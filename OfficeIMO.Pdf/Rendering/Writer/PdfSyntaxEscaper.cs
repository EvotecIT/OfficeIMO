using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfSyntaxEscaper {
    internal static string IndirectReference(int objectNumber, int generation = 0) {
        if (objectNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(objectNumber), "PDF object number must be positive.");
        }

        if (generation < 0) {
            throw new ArgumentOutOfRangeException(nameof(generation), "PDF generation number cannot be negative.");
        }

        return objectNumber.ToString(CultureInfo.InvariantCulture) +
            " " +
            generation.ToString(CultureInfo.InvariantCulture) +
            " R";
    }

    internal static string LiteralString(string value) {
        return "(" + EscapeLiteralContent(value) + ")";
    }

    internal static string EscapeLiteralContent(string value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        var sb = new StringBuilder(value.Length + 8);
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            switch (ch) {
                case '\\': sb.Append("\\\\"); break;
                case '(': sb.Append("\\("); break;
                case ')': sb.Append("\\)"); break;
                case '\r': sb.Append("\\r"); break;
                case '\n': sb.Append("\\n"); break;
                case '\t': sb.Append("\\t"); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                default:
                    if (ch < 32 || ch == 127) {
                        int v = ch;
                        sb.Append('\\')
                            .Append(((v >> 6) & 0x7).ToString(CultureInfo.InvariantCulture))
                            .Append(((v >> 3) & 0x7).ToString(CultureInfo.InvariantCulture))
                            .Append((v & 0x7).ToString(CultureInfo.InvariantCulture));
                    } else {
                        sb.Append(ch);
                    }

                    break;
            }
        }

        return sb.ToString();
    }

    internal static string Name(string value) {
        Guard.NotNull(value, nameof(value));
        var sb = new StringBuilder(value.Length);
        foreach (char ch in value) {
            if (ch <= 0x20 || ch >= 0x7F || IsNameDelimiter(ch)) {
                sb.Append('#').Append(((int)ch).ToString("X2", CultureInfo.InvariantCulture));
            } else {
                sb.Append(ch);
            }
        }

        return sb.ToString();
    }

    private static bool IsNameDelimiter(char ch) {
        switch (ch) {
            case '(':
            case ')':
            case '<':
            case '>':
            case '[':
            case ']':
            case '{':
            case '}':
            case '/':
            case '%':
            case '#':
                return true;
            default:
                return false;
        }
    }
}
