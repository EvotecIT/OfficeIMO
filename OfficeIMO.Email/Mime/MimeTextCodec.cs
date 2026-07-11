using System.Text.RegularExpressions;

namespace OfficeIMO.Email;

internal static class MimeTextCodec {
    private static readonly Regex EncodedWordPattern = new Regex(
        @"=\?([^?\s]+)\?([bBqQ])\?([^?]*)\?=",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex EncodedWordSeparatorPattern = new Regex(
        @"(?<=\?=)[ \t\r\n]+(?==\?)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    static MimeTextCodec() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    internal static string DecodeHeader(string value, IList<EmailDiagnostic> diagnostics, string location) {
        if (string.IsNullOrEmpty(value) || value.IndexOf("=?", StringComparison.Ordinal) < 0) return value;

        try {
            string adjacentWords = EncodedWordSeparatorPattern.Replace(value, string.Empty);
            return EncodedWordPattern.Replace(adjacentWords, match => {
                string charset = match.Groups[1].Value;
                string encoding = match.Groups[2].Value;
                string payload = match.Groups[3].Value;
                byte[] bytes = string.Equals(encoding, "B", StringComparison.OrdinalIgnoreCase)
                    ? DecodeBase64(payload, diagnostics, location)
                    : DecodeQuotedPrintable(Encoding.ASCII.GetBytes(payload.Replace('_', ' ')), true, diagnostics, location);
                return DecodeText(bytes, charset, diagnostics, location);
            });
        } catch (Exception ex) when (ex is FormatException || ex is ArgumentException) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_ENCODED_WORD_INVALID", ex.Message, EmailDiagnosticSeverity.Warning, location));
            return value;
        }
    }

    internal static string DecodeText(byte[] bytes, string? charset, IList<EmailDiagnostic> diagnostics, string location) {
        string normalized = (charset ?? "utf-8").Trim().Trim('"').ToLowerInvariant();
        try {
            switch (normalized) {
                case "us-ascii":
                case "ascii":
                    return Encoding.ASCII.GetString(bytes);
                case "utf-8":
                case "utf8":
                    return new UTF8Encoding(false, false).GetString(bytes);
                case "utf-16":
                case "unicode":
                    return Encoding.Unicode.GetString(bytes);
                case "utf-16be":
                    return Encoding.BigEndianUnicode.GetString(bytes);
                case "iso-8859-1":
                case "latin1":
                case "latin-1":
                    return DecodeLatin1(bytes);
                case "windows-1252":
                case "cp1252":
                    return DecodeWindows1252(bytes);
                default:
                    return Encoding.GetEncoding(normalized).GetString(bytes);
            }
        } catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_CHARSET_UNSUPPORTED",
                string.Concat("Charset '", normalized, "' is unavailable; UTF-8 fallback was used."),
                EmailDiagnosticSeverity.Warning, location));
            return new UTF8Encoding(false, false).GetString(bytes);
        }
    }

    internal static byte[] DecodeTransfer(byte[] bytes, string? transferEncoding, IList<EmailDiagnostic> diagnostics, string location) {
        string normalized = (transferEncoding ?? string.Empty).Trim().ToLowerInvariant();
        switch (normalized) {
            case "base64":
                return DecodeBase64(Encoding.ASCII.GetString(bytes), diagnostics, location);
            case "quoted-printable":
                return DecodeQuotedPrintable(bytes, false, diagnostics, location);
            case "7bit":
            case "8bit":
            case "binary":
            case "":
                return bytes;
            default:
                diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_TRANSFER_ENCODING_UNKNOWN",
                    string.Concat("Transfer encoding '", normalized, "' was preserved without decoding."),
                    EmailDiagnosticSeverity.Warning, location));
                return bytes;
        }
    }

    internal static byte[] DecodeBase64(string value, IList<EmailDiagnostic> diagnostics, string location) {
        string compact = RemoveWhiteSpace(value);
        if (compact.Length == 0) return Array.Empty<byte>();
        try {
            return Convert.FromBase64String(compact);
        } catch (FormatException) {
            int remainder = compact.Length % 4;
            if (remainder != 0) compact = compact.PadRight(compact.Length + (4 - remainder), '=');
            try {
                byte[] recovered = Convert.FromBase64String(compact);
                diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BASE64_PADDING_RECOVERED",
                    "Missing Base64 padding was recovered.", EmailDiagnosticSeverity.Warning, location));
                return recovered;
            } catch (FormatException ex) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BASE64_INVALID", ex.Message, EmailDiagnosticSeverity.Error, location));
                return Encoding.ASCII.GetBytes(value);
            }
        }
    }

    internal static byte[] DecodeQuotedPrintable(byte[] input, bool headerMode, IList<EmailDiagnostic> diagnostics, string location) {
        using (MemoryStream output = new MemoryStream(input.Length)) {
            for (int i = 0; i < input.Length; i++) {
                byte current = input[i];
                if (current != '=') {
                    output.WriteByte(current);
                    continue;
                }

                if (i + 1 < input.Length && input[i + 1] == '\n') {
                    i++;
                    continue;
                }
                if (i + 2 < input.Length && input[i + 1] == '\r' && input[i + 2] == '\n') {
                    i += 2;
                    continue;
                }
                if (i + 2 < input.Length && TryHex(input[i + 1], out int high) && TryHex(input[i + 2], out int low)) {
                    output.WriteByte((byte)((high << 4) | low));
                    i += 2;
                    continue;
                }

                output.WriteByte(current);
                if (!headerMode) {
                    diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_QUOTED_PRINTABLE_INVALID",
                        "An invalid quoted-printable escape was preserved.", EmailDiagnosticSeverity.Warning, location));
                }
            }
            return output.ToArray();
        }
    }

    private static string RemoveWhiteSpace(string value) {
        StringBuilder builder = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            if (!char.IsWhiteSpace(value[i])) builder.Append(value[i]);
        }
        return builder.ToString();
    }

    private static bool TryHex(byte value, out int result) {
        if (value >= '0' && value <= '9') {
            result = value - '0';
            return true;
        }
        if (value >= 'A' && value <= 'F') {
            result = value - 'A' + 10;
            return true;
        }
        if (value >= 'a' && value <= 'f') {
            result = value - 'a' + 10;
            return true;
        }
        result = 0;
        return false;
    }

    private static string DecodeLatin1(byte[] bytes) {
        char[] characters = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) characters[i] = (char)bytes[i];
        return new string(characters);
    }

    private static string DecodeWindows1252(byte[] bytes) {
        const string replacements = "\u20AC\u0081\u201A\u0192\u201E\u2026\u2020\u2021\u02C6\u2030\u0160\u2039\u0152\u008D\u017D\u008F" +
            "\u0090\u2018\u2019\u201C\u201D\u2022\u2013\u2014\u02DC\u2122\u0161\u203A\u0153\u009D\u017E\u0178";
        char[] characters = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) {
            byte value = bytes[i];
            characters[i] = value >= 0x80 && value <= 0x9F ? replacements[value - 0x80] : (char)value;
        }
        return new string(characters);
    }
}
