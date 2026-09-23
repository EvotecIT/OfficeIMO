using System.Globalization;
using System.Threading;

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

    /// <summary>Appends an indirect reference without allocating the combined reference string.</summary>
    internal static void AppendIndirectReference(StringBuilder destination, int objectNumber, int generation = 0) {
        Guard.NotNull(destination, nameof(destination));
        if (objectNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(objectNumber), "PDF object number must be positive.");
        }

        if (generation < 0) {
            throw new ArgumentOutOfRangeException(nameof(generation), "PDF generation number cannot be negative.");
        }

#if NET6_0_OR_GREATER
        Span<char> buffer = stackalloc char[11];
        if (!objectNumber.TryFormat(buffer, out int written, default, CultureInfo.InvariantCulture)) {
            throw new InvalidOperationException("The PDF object number could not be formatted.");
        }
        destination.Append(buffer.Slice(0, written)).Append(' ');
        if (!generation.TryFormat(buffer, out written, default, CultureInfo.InvariantCulture)) {
            throw new InvalidOperationException("The PDF generation number could not be formatted.");
        }
        destination.Append(buffer.Slice(0, written)).Append(" R");
#else
        destination.Append(objectNumber.ToString(CultureInfo.InvariantCulture))
            .Append(' ')
            .Append(generation.ToString(CultureInfo.InvariantCulture))
            .Append(" R");
#endif
    }

    internal static string LiteralString(string value, CancellationToken cancellationToken = default) {
        Guard.NotNull(value, nameof(value));
        cancellationToken.ThrowIfCancellationRequested();
        for (int index = 0; index < value.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (value[index] > byte.MaxValue) {
                return TextString(value, cancellationToken);
            }
        }

        return "(" + EscapeLiteralContent(value, cancellationToken) + ")";
    }

    internal static string WinAnsiHexString(string value, CancellationToken cancellationToken = default) {
        Guard.NotNull(value, nameof(value));
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = PdfWinAnsiEncoding.Encode(value, cancellationToken);
        return HexString(bytes, cancellationToken);
    }

    internal static string TextString(string value, CancellationToken cancellationToken = default) {
        Guard.NotNull(value, nameof(value));
        cancellationToken.ThrowIfCancellationRequested();
        if (PdfWinAnsiEncoding.CanEncode(value, out _, cancellationToken)) {
            return WinAnsiHexString(value, cancellationToken);
        }

        byte[] bytes = new byte[2 + value.Length * 2];
        bytes[0] = 0xFE;
        bytes[1] = 0xFF;
        for (int i = 0; i < value.Length; i++) {
            if ((i & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            char ch = value[i];
            bytes[2 + i * 2] = (byte)(ch >> 8);
            bytes[3 + i * 2] = (byte)(ch & 0xFF);
        }

        return HexString(bytes, cancellationToken);
    }

    internal static void AppendTextStringCancellable(StringBuilder destination, string value, CancellationToken cancellationToken) {
        Guard.NotNull(destination, nameof(destination));
        Guard.NotNull(value, nameof(value));
        cancellationToken.ThrowIfCancellationRequested();
        destination.Append('<');
        if (PdfWinAnsiEncoding.CanEncode(value, out _, cancellationToken)) {
            byte[] bytes = PdfWinAnsiEncoding.Encode(value, cancellationToken);
            for (int index = 0; index < bytes.Length; index++) {
                if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                AppendHexByte(destination, bytes[index]);
            }
        } else {
            destination.Append("FEFF");
            for (int index = 0; index < value.Length; index++) {
                if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                AppendHexByte(destination, (byte)(value[index] >> 8));
                AppendHexByte(destination, (byte)value[index]);
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        destination.Append('>');
    }

    private static void AppendHexByte(StringBuilder destination, byte value) {
        const string digits = "0123456789ABCDEF";
        destination.Append(digits[value >> 4]).Append(digits[value & 0x0F]);
    }

    internal static string HexString(byte[] bytes, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var sb = new StringBuilder(bytes.Length * 2 + 2);
        sb.Append('<');
        for (int i = 0; i < bytes.Length; i++) {
            if ((i & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            sb.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        }

        sb.Append('>');
        return sb.ToString();
    }

    internal static string EscapeLiteralContent(string value, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        var sb = new StringBuilder(value.Length + 8);
        for (int i = 0; i < value.Length; i++) {
            if ((i & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
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
        AppendName(sb, value);
        return sb.ToString();
    }

    /// <summary>Appends an escaped PDF name without allocating an intermediate escaped string.</summary>
    internal static void AppendName(StringBuilder destination, string value, CancellationToken cancellationToken = default) {
        Guard.NotNull(destination, nameof(destination));
        Guard.NotNull(value, nameof(value));
        cancellationToken.ThrowIfCancellationRequested();

        bool requiresUtf8 = false;
        for (int index = 0; index < value.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (value[index] >= 0x80) {
                requiresUtf8 = true;
                break;
            }
        }

        if (requiresUtf8) {
            for (int index = 0; index < value.Length;) {
                cancellationToken.ThrowIfCancellationRequested();
                int length = Math.Min(1024, value.Length - index);
                if (index + length < value.Length && char.IsHighSurrogate(value[index + length - 1]) && char.IsLowSurrogate(value[index + length])) length++;
                byte[] bytes = Encoding.UTF8.GetBytes(value.Substring(index, length));
                foreach (byte encodedByte in bytes) AppendNameByte(destination, encodedByte);
                index += length;
            }
            return;
        }

        for (int index = 0; index < value.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            AppendNameByte(destination, (byte)value[index]);
        }
    }

    private static void AppendNameByte(StringBuilder destination, byte value) {
        char ch = (char)value;
        if (value <= 0x20 || value >= 0x7F || IsNameDelimiter(ch)) {
            const string HexDigits = "0123456789ABCDEF";
            destination.Append('#')
                .Append(HexDigits[value >> 4])
                .Append(HexDigits[value & 0x0F]);
        } else {
            destination.Append(ch);
        }
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
