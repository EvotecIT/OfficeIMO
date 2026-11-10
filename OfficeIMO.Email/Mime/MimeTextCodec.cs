using OfficeIMO.Core.Internal;
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

    internal static bool ContainsEncodedWord(string value) =>
        !string.IsNullOrEmpty(value) && EncodedWordPattern.IsMatch(value);

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
        if (string.IsNullOrWhiteSpace(charset)) {
            try {
                return new UTF8Encoding(false, true).GetString(bytes);
            } catch (DecoderFallbackException) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_CHARSET_GUESSED",
                    "The charsetless text was not valid UTF-8; Windows-1252 recovery was used.",
                    EmailDiagnosticSeverity.Warning, location));
                return DecodeWindows1252(bytes);
            }
        }

        string normalized = charset!.Trim().Trim('"').ToLowerInvariant();
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

    internal static bool TryDecodeTextStrict(byte[] bytes, string? charset, out string value) {
        value = string.Empty;
        try {
            value = ResolveStrictEncoding(charset).GetString(bytes);
            return true;
        } catch (Exception exception) when (exception is ArgumentException
                                            || exception is NotSupportedException
                                            || exception is DecoderFallbackException) {
            return false;
        }
    }

    internal static byte[] EncodeText(string text, string? charset) {
        ArgumentNullException.ThrowIfNull(text);
        return ResolveStrictEncoding(charset).GetBytes(text);
    }

    internal static Encoding ResolveStrictEncoding(string? charset) {
        string normalized = string.IsNullOrWhiteSpace(charset)
            ? "utf-8"
            : charset!.Trim().Trim('"').ToLowerInvariant();
        Encoding encoding;
        switch (normalized) {
            case "us-ascii":
            case "ascii":
                encoding = Encoding.GetEncoding(Encoding.ASCII.CodePage,
                    EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
                break;
            case "utf-8":
            case "utf8":
                encoding = new UTF8Encoding(false, true);
                break;
            case "utf-16":
            case "unicode":
                encoding = new UnicodeEncoding(false, false, true);
                break;
            case "utf-16be":
                encoding = new UnicodeEncoding(true, false, true);
                break;
            case "iso-8859-1":
            case "latin1":
            case "latin-1":
                encoding = Encoding.GetEncoding(28591,
                    EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
                break;
            case "windows-1252":
            case "cp1252":
                encoding = Encoding.GetEncoding(1252,
                    EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
                break;
            default:
                encoding = Encoding.GetEncoding(normalized,
                    EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
                break;
        }
        return encoding;
    }

    internal static string DecodeText(byte[] bytes, int codePage, IList<EmailDiagnostic> diagnostics,
        string location) {
        try {
            return Encoding.GetEncoding(codePage).GetString(bytes);
        } catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_CHARSET_UNSUPPORTED",
                string.Concat("Code page ", codePage.ToString(CultureInfo.InvariantCulture),
                    " is unavailable; UTF-8 fallback was used."),
                EmailDiagnosticSeverity.Warning, location));
            return new UTF8Encoding(false, false).GetString(bytes);
        }
    }

    internal static byte[] DecodeTransfer(byte[] bytes, string? transferEncoding, IList<EmailDiagnostic> diagnostics, string location) {
        long decodedLength = GetDecodedLength(bytes, 0, bytes.Length, transferEncoding);
        return DecodeTransfer(bytes, 0, bytes.Length, decodedLength, transferEncoding, diagnostics, location);
    }

    internal static bool IsSupportedTransferEncoding(string? transferEncoding) {
        string normalized = (transferEncoding ?? string.Empty).Trim().ToLowerInvariant();
        return normalized == ""
            || normalized == "7bit"
            || normalized == "8bit"
            || normalized == "binary"
            || normalized == "base64"
            || normalized == "quoted-printable";
    }

    internal static byte[] DecodeTransfer(byte[] bytes, int offset, int count, long decodedLength,
        string? transferEncoding, IList<EmailDiagnostic> diagnostics, string location) {
        string normalized = (transferEncoding ?? string.Empty).Trim().ToLowerInvariant();
        switch (normalized) {
            case "base64":
                return DecodeBase64(bytes, offset, count, decodedLength, diagnostics, location);
            case "quoted-printable":
                return DecodeQuotedPrintable(bytes, offset, count, checked((int)decodedLength), false, diagnostics, location);
            case "7bit":
            case "8bit":
            case "binary":
            case "":
                return Copy(bytes, offset, count);
            default:
                diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_TRANSFER_ENCODING_UNKNOWN",
                    string.Concat("Transfer encoding '", normalized, "' was preserved without decoding."),
                    EmailDiagnosticSeverity.Warning, location));
                return Copy(bytes, offset, count);
        }
    }

    internal static long GetDecodedLength(byte[] bytes, int offset, int count, string? transferEncoding,
        IList<EmailDiagnostic>? diagnostics = null, string? location = null) {
        string normalized = (transferEncoding ?? string.Empty).Trim().ToLowerInvariant();
        switch (normalized) {
            case "base64":
                long base64Length = GetBase64DecodedLength(bytes, offset, count,
                    out bool validBase64, out bool recoveredPadding);
                if (diagnostics != null && !validBase64) {
                    diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BASE64_INVALID",
                        "The invalid Base64 payload was preserved without decoding.",
                        EmailDiagnosticSeverity.Error, location));
                } else if (diagnostics != null && recoveredPadding) {
                    diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BASE64_PADDING_RECOVERED",
                        "Missing Base64 padding was recovered.", EmailDiagnosticSeverity.Warning, location));
                }
                return base64Length;
            case "quoted-printable":
                long quotedLength = GetQuotedPrintableDecodedLength(bytes, offset, count, out bool invalidQuoted);
                if (diagnostics != null && invalidQuoted) {
                    diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_QUOTED_PRINTABLE_INVALID",
                        "An invalid quoted-printable escape was preserved.",
                        EmailDiagnosticSeverity.Warning, location));
                }
                return quotedLength;
            case "7bit":
            case "8bit":
            case "binary":
            case "":
                return count;
            default:
                diagnostics?.Add(new EmailDiagnostic("EMAIL_MIME_TRANSFER_ENCODING_UNKNOWN",
                    string.Concat("Transfer encoding '", normalized, "' was preserved without decoding."),
                    EmailDiagnosticSeverity.Warning, location));
                return count;
        }
    }

    private static long GetBase64DecodedLength(byte[] bytes, int offset, int count,
        out bool valid, out bool recoveredPadding) =>
        OfficeBase64.GetMimeDecodedLength(bytes, offset, count, out valid, out recoveredPadding);

    private static long GetQuotedPrintableDecodedLength(byte[] bytes, int offset, int count, out bool invalid) {
        long length = count;
        int end = checked(offset + count);
        invalid = false;
        for (int index = offset; index < end; index++) {
            index = Array.IndexOf(bytes, (byte)'=', index, end - index);
            if (index < 0) break;
            if (index + 1 < end && bytes[index + 1] == '\n') {
                length -= 2;
                index++;
            } else if (index + 2 < end && bytes[index + 1] == '\r' && bytes[index + 2] == '\n') {
                length -= 3;
                index += 2;
            } else if (index + 2 < end && TryHex(bytes[index + 1], out _) && TryHex(bytes[index + 2], out _)) {
                length -= 2;
                index += 2;
            } else {
                invalid = true;
            }
        }
        return length;
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
                diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BASE64_INVALID", ex.Message,
                    EmailDiagnosticSeverity.Error, location));
                return Encoding.ASCII.GetBytes(value);
            }
        }
    }

    private static byte[] DecodeBase64(byte[] input, int offset, int count, long decodedLength,
        IList<EmailDiagnostic> diagnostics, string location) {
        if (decodedLength == 0) return Array.Empty<byte>();
        byte[] output = new byte[checked((int)decodedLength)];
        if (OfficeBase64.TryDecodeMime(input, offset, count, output, out int written, out bool recoveredPadding)
            && written == output.Length) {
            if (recoveredPadding) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BASE64_PADDING_RECOVERED",
                    "Missing Base64 padding was recovered.", EmailDiagnosticSeverity.Warning, location));
            }
            return output;
        }
        diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BASE64_INVALID",
            "The invalid Base64 payload was preserved without decoding.",
            EmailDiagnosticSeverity.Error, location));
        if (output.Length != count) output = new byte[count];
        Buffer.BlockCopy(input, offset, output, 0, count);
        return output;
    }

    internal static byte[] DecodeQuotedPrintable(byte[] input, bool headerMode,
        IList<EmailDiagnostic> diagnostics, string location) =>
        DecodeQuotedPrintable(input, 0, input.Length,
            checked((int)GetQuotedPrintableDecodedLength(input, 0, input.Length, out _)),
            headerMode, diagnostics, location);

    private static byte[] DecodeQuotedPrintable(byte[] input, int offset, int count, int decodedLength,
        bool headerMode, IList<EmailDiagnostic> diagnostics, string location) {
        var output = new byte[decodedLength];
        int written = 0;
        int end = checked(offset + count);
        while (offset < end) {
            int equals = Array.IndexOf(input, (byte)'=', offset, end - offset);
            if (equals < 0) equals = end;
            int run = equals - offset;
            Buffer.BlockCopy(input, offset, output, written, run);
            written += run;
            offset = equals;
            if (offset == end) break;
            if (offset + 1 < end && input[offset + 1] == '\n') {
                offset += 2;
            } else if (offset + 2 < end && input[offset + 1] == '\r' && input[offset + 2] == '\n') {
                offset += 3;
            } else if (offset + 2 < end && TryHex(input[offset + 1], out int high) && TryHex(input[offset + 2], out int low)) {
                output[written++] = (byte)(high << 4 | low);
                offset += 3;
            } else {
                output[written++] = input[offset++];
                if (!headerMode) {
                    diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_QUOTED_PRINTABLE_INVALID",
                        "An invalid quoted-printable escape was preserved.", EmailDiagnosticSeverity.Warning, location));
                }
            }
        }
        return output;
    }

    private static string RemoveWhiteSpace(string value) {
        StringBuilder builder = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            if (!char.IsWhiteSpace(value[i])) builder.Append(value[i]);
        }
        return builder.ToString();
    }

    private static byte[] Copy(byte[] input, int offset, int count) {
        if (offset == 0 && count == input.Length) return input;
        byte[] result = new byte[count];
        Buffer.BlockCopy(input, offset, result, 0, count);
        return result;
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
