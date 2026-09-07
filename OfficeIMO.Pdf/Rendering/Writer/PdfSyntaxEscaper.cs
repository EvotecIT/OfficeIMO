using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfSyntaxEscaper {
    /// <summary>Writes a finite real without precision truncation or PDF-invalid exponent notation.</summary>
    internal static string Number(double value) {
        if (double.IsNaN(value) || double.IsInfinity(value))
            throw new ArgumentOutOfRangeException(nameof(value), "PDF numbers must be finite.");
        if (value == 0D) return "0";
        string text = value.ToString("R", CultureInfo.InvariantCulture);
        int exponentIndex = text.IndexOf('E');
        if (exponentIndex < 0) return text;
#if NETSTANDARD2_0 || NETFRAMEWORK
        int exponent = int.Parse(text.Substring(exponentIndex + 1), CultureInfo.InvariantCulture);
#else
        int exponent = int.Parse(text.AsSpan(exponentIndex + 1), CultureInfo.InvariantCulture);
#endif
        string mantissa = text.Substring(0, exponentIndex);
        string sign = mantissa[0] == '-' ? "-" : string.Empty;
        if (sign.Length > 0) mantissa = mantissa.Substring(1);
        int point = mantissa.IndexOf('.');
        int decimalPosition = (point < 0 ? mantissa.Length : point) + exponent;
        string digits = mantissa.Replace(".", string.Empty);
        if (decimalPosition <= 0) return sign + "0." + new string('0', -decimalPosition) + digits;
        if (decimalPosition >= digits.Length) return sign + digits + new string('0', decimalPosition - digits.Length);
        return sign + digits.Insert(decimalPosition, ".");
    }

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
        Guard.NotNull(value, nameof(value));
        for (int index = 0; index < value.Length; index++) {
            if (value[index] > byte.MaxValue) {
                return TextString(value);
            }
        }

        return "(" + EscapeLiteralContent(value) + ")";
    }

    internal static string WinAnsiHexString(string value) {
        Guard.NotNull(value, nameof(value));
        byte[] bytes = PdfWinAnsiEncoding.Encode(value);
        return HexString(bytes);
    }

    internal static string TextString(string value) {
        Guard.NotNull(value, nameof(value));
        if (PdfWinAnsiEncoding.CanEncode(value, out _)) {
            return WinAnsiHexString(value);
        }

        byte[] bytes = new byte[2 + value.Length * 2];
        bytes[0] = 0xFE;
        bytes[1] = 0xFF;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            bytes[2 + i * 2] = (byte)(ch >> 8);
            bytes[3 + i * 2] = (byte)(ch & 0xFF);
        }

        return HexString(bytes);
    }

    internal static string HexString(byte[] bytes) {
        var sb = new StringBuilder(bytes.Length * 2 + 2);
        sb.Append('<');
        for (int i = 0; i < bytes.Length; i++) {
            sb.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        }

        sb.Append('>');
        return sb.ToString();
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
