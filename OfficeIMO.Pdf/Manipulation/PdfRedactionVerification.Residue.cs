using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionVerification {
    private const int MaxDecodedRedactionVerificationStreamBytes = 16 * 1024 * 1024;

    private static bool ContainsEncodedPdfMarker(byte[] pdf, string marker) {
        if (string.IsNullOrEmpty(marker)) {
            return false;
        }

        byte[][] encodings = BuildMarkerEncodings(marker);
        for (int i = 0; i < encodings.Length; i++) {
            if (ContainsBytes(pdf, encodings[i]) ||
                ContainsLiteralStringBytes(pdf, encodings[i]) ||
                ContainsHexStringBytes(pdf, encodings[i])) {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsDecodedStreamMarker(byte[] pdf, string marker) {
        if (string.IsNullOrEmpty(marker)) {
            return false;
        }

        Dictionary<int, PdfIndirectObject> objects;
        try {
            objects = PdfSyntax.ParseObjects(pdf).Map;
        } catch (Exception ex) when (ex is not OutOfMemoryException) {
            return false;
        }

        byte[][] encodings = BuildMarkerEncodings(marker);
        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is not PdfStream stream || stream.DecodingFailed) {
                continue;
            }

            if (!StreamDecoder.TryDecode(stream.Dictionary, stream.Data, MaxDecodedRedactionVerificationStreamBytes, out byte[] decoded, objects)) {
                continue;
            }

            for (int i = 0; i < encodings.Length; i++) {
                if (ContainsBytes(decoded, encodings[i]) ||
                    ContainsLiteralStringBytes(decoded, encodings[i]) ||
                    ContainsHexStringBytes(decoded, encodings[i])) {
                    return true;
                }
            }
        }

        return false;
    }

    private static List<PdfRedactionVerificationIssue> FindUndecodableStreamIssues(byte[] pdf) {
        var issues = new List<PdfRedactionVerificationIssue>();
        Dictionary<int, PdfIndirectObject> objects;
        try {
            objects = PdfSyntax.ParseObjects(pdf).Map;
        } catch (Exception ex) when (ex is not OutOfMemoryException) {
            issues.Add(new PdfRedactionVerificationIssue(
                "DecodedPdfStreamInspection",
                "PDF",
                "PDF streams could not be inspected during redaction verification: " + ex.GetType().Name));
            return issues;
        }

        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is not PdfStream stream) {
                continue;
            }

            string objectReference = indirect.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (stream.DecodingFailed) {
                issues.Add(CreateUndecodableStreamIssue(objectReference, stream.DecodingError ?? "PDF parser reported a stream decoding failure."));
                continue;
            }

            if (!StreamDecoder.TryDecode(stream.Dictionary, stream.Data, MaxDecodedRedactionVerificationStreamBytes, out _, objects)) {
                issues.Add(CreateUndecodableStreamIssue(
                    objectReference,
                    "The stream filter is unsupported, uses active decode parameters, exceeds the verification size limit, or failed decoding."));
            }
        }

        return issues;
    }

    private static PdfRedactionVerificationIssue CreateUndecodableStreamIssue(string objectReference, string reason) {
        return new PdfRedactionVerificationIssue(
            "UndecodablePdfStream",
            objectReference,
            "PDF stream object " + objectReference + " could not be decoded during redaction verification; hidden removed content cannot be ruled out. " + reason);
    }

    private static byte[][] BuildMarkerEncodings(string marker) {
        var encodings = new List<byte[]>();
        AddEncoding(encodings, PdfEncoding.Latin1GetBytes(marker));
        AddEncoding(encodings, System.Text.Encoding.UTF8.GetBytes(marker));
        AddEncoding(encodings, System.Text.Encoding.BigEndianUnicode.GetBytes(marker));
        AddEncoding(encodings, System.Text.Encoding.Unicode.GetBytes(marker));
        return encodings.ToArray();
    }

    private static void AddEncoding(List<byte[]> encodings, byte[] candidate) {
        for (int i = 0; i < encodings.Count; i++) {
            if (BytesEqual(encodings[i], candidate)) {
                return;
            }
        }

        encodings.Add(candidate);
    }

    private static bool ContainsBytes(byte[] haystack, byte[] needle) {
        if (needle.Length == 0 || needle.Length > haystack.Length) {
            return false;
        }

        for (int i = 0; i <= haystack.Length - needle.Length; i++) {
            int j = 0;
            while (j < needle.Length && haystack[i + j] == needle[j]) {
                j++;
            }

            if (j == needle.Length) {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsLiteralStringBytes(byte[] pdf, byte[] markerBytes) {
        for (int i = 0; i < pdf.Length; i++) {
            if (pdf[i] != (byte)'(') {
                continue;
            }

            if (TryReadLiteralStringBytes(pdf, i, out byte[] literalBytes, out int end)) {
                if (ContainsBytes(literalBytes, markerBytes)) {
                    return true;
                }

                i = end;
            }
        }

        return false;
    }

    private static bool TryReadLiteralStringBytes(byte[] pdf, int start, out byte[] literalBytes, out int end) {
        var bytes = new List<byte>();
        int depth = 1;
        for (int i = start + 1; i < pdf.Length; i++) {
            byte value = pdf[i];
            if (value == (byte)'\\') {
                if (i + 1 >= pdf.Length) {
                    break;
                }

                i++;
                byte escaped = pdf[i];
                if (TryReadOctalEscape(pdf, i, out byte octalValue, out int octalEnd)) {
                    bytes.Add(octalValue);
                    i = octalEnd;
                    continue;
                }

                if (TryMapLiteralEscape(escaped, out byte mapped)) {
                    bytes.Add(mapped);
                    continue;
                }

                if (IsLineBreak(escaped)) {
                    if (escaped == (byte)'\r' && i + 1 < pdf.Length && pdf[i + 1] == (byte)'\n') {
                        i++;
                    }

                    continue;
                }

                bytes.Add(escaped);
                continue;
            }

            if (value == (byte)'(') {
                depth++;
                bytes.Add(value);
                continue;
            }

            if (value == (byte)')') {
                depth--;
                if (depth == 0) {
                    literalBytes = bytes.ToArray();
                    end = i;
                    return true;
                }

                bytes.Add(value);
                continue;
            }

            bytes.Add(value);
        }

        literalBytes = Array.Empty<byte>();
        end = start;
        return false;
    }

    private static bool TryReadOctalEscape(byte[] pdf, int firstDigit, out byte value, out int end) {
        value = 0;
        end = firstDigit;
        if (!IsOctalDigit(pdf[firstDigit])) {
            return false;
        }

        int result = 0;
        int count = 0;
        int index = firstDigit;
        while (index < pdf.Length && count < 3 && IsOctalDigit(pdf[index])) {
            result = (result * 8) + (pdf[index] - (byte)'0');
            index++;
            count++;
        }

        value = (byte)(result & 0xFF);
        end = index - 1;
        return true;
    }

    private static bool TryMapLiteralEscape(byte escaped, out byte mapped) {
        switch (escaped) {
            case (byte)'n':
                mapped = (byte)'\n';
                return true;
            case (byte)'r':
                mapped = (byte)'\r';
                return true;
            case (byte)'t':
                mapped = (byte)'\t';
                return true;
            case (byte)'b':
                mapped = 8;
                return true;
            case (byte)'f':
                mapped = 12;
                return true;
            case (byte)'(':
            case (byte)')':
            case (byte)'\\':
                mapped = escaped;
                return true;
            default:
                mapped = escaped;
                return false;
        }
    }

    private static bool ContainsHexStringBytes(byte[] pdf, byte[] markerBytes) {
        string markerHex = ToUpperHex(markerBytes);
        for (int i = 0; i < pdf.Length; i++) {
            if (pdf[i] != (byte)'<' || (i + 1 < pdf.Length && pdf[i + 1] == (byte)'<')) {
                continue;
            }

            int end = FindHexStringEnd(pdf, i + 1);
            if (end < 0) {
                continue;
            }

            if (HexStringContains(pdf, i + 1, end, markerHex)) {
                return true;
            }

            i = end;
        }

        return false;
    }

    private static int FindHexStringEnd(byte[] pdf, int start) {
        for (int i = start; i < pdf.Length; i++) {
            if (pdf[i] == (byte)'>') {
                return i;
            }

            if (!IsHexDigit(pdf[i]) && !IsPdfWhiteSpace(pdf[i])) {
                return -1;
            }
        }

        return -1;
    }

    private static bool HexStringContains(byte[] pdf, int start, int end, string markerHex) {
        var builder = new System.Text.StringBuilder(end - start);
        for (int i = start; i < end; i++) {
            byte value = pdf[i];
            if (IsHexDigit(value)) {
                builder.Append((char)ToUpperAscii(value));
            }
        }

        return builder.ToString().Contains(markerHex);
    }

    private static string ToUpperHex(byte[] bytes) {
        var builder = new System.Text.StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(bytes[i].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static bool BytesEqual(byte[] left, byte[] right) {
        if (left.Length != right.Length) {
            return false;
        }

        for (int i = 0; i < left.Length; i++) {
            if (left[i] != right[i]) {
                return false;
            }
        }

        return true;
    }

    private static bool IsHexDigit(byte value) {
        return (value >= (byte)'0' && value <= (byte)'9') ||
            (value >= (byte)'A' && value <= (byte)'F') ||
            (value >= (byte)'a' && value <= (byte)'f');
    }

    private static bool IsOctalDigit(byte value) {
        return value >= (byte)'0' && value <= (byte)'7';
    }

    private static bool IsLineBreak(byte value) {
        return value == (byte)'\n' || value == (byte)'\r';
    }

    private static bool IsPdfWhiteSpace(byte value) {
        return value == 0 ||
            value == 9 ||
            value == 10 ||
            value == 12 ||
            value == 13 ||
            value == 32;
    }

    private static byte ToUpperAscii(byte value) {
        if (value >= (byte)'a' && value <= (byte)'f') {
            return (byte)(value - 32);
        }

        return value;
    }
}
