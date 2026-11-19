using System.Text;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private const int PdfStartXrefTailBytes = 2048;
    private const int PdfTrailerWindowBytes = 64 * 1024;
    private const int PdfXrefSubsectionHeaderBytes = 128;
    private const int MaximumPdfXrefSubsections = 4096;

    /// <summary>
    /// Reports whether the trailer dictionary, or the cross-reference stream dictionary that replaces it, has an
    /// <c>Encrypt</c> entry. The prefix is scanned from the start of the file, which covers the first-page trailer of a
    /// linearized file; for a seekable stream the cross-reference section named by the final <c>startxref</c> is scanned
    /// too. Every trailer of an encrypted file carries the entry.
    /// </summary>
    private static bool HasPdfEncryptDictionary(Stream stream, long start, bool seekable, byte[] prefix,
        CancellationToken cancellationToken) {
        if (ContainsPdfTrailerEncryptEntry(prefix, cancellationToken)) return true;
        if (!seekable) return false;

        long originalPosition = stream.Position;
        try {
            return HasPdfEncryptDictionaryAtStartXref(stream, start, prefix, cancellationToken);
        } catch (Exception exception) when (exception is IOException or ArgumentException or OverflowException) {
            // Detection reports what it can prove; a malformed cross-reference section is not a detection failure.
            return false;
        } finally {
            stream.Position = originalPosition;
        }
    }

    private static bool HasPdfEncryptDictionaryAtStartXref(Stream stream, long start, byte[] prefix,
        CancellationToken cancellationToken) {
        long length = stream.Length - start;
        if (length <= prefix.Length) return false;
        int tailBytes = (int)Math.Min(PdfStartXrefTailBytes, length);
        byte[] tail = ReadPdfWindow(stream, start, length - tailBytes, tailBytes);
        if (!TryReadPdfStartXref(tail, out long crossReferenceOffset) || crossReferenceOffset >= length) return false;
        long trailerOffset = LocatePdfTrailerDictionary(stream, start, crossReferenceOffset, length, cancellationToken);
        if (trailerOffset < 0 || trailerOffset >= length) return false;
        byte[] window = ReadPdfWindow(stream, start, trailerOffset, (int)Math.Min(PdfTrailerWindowBytes, length - trailerOffset));
        return ContainsPdfTrailerEncryptEntry(window, cancellationToken);
    }

    /// <summary>
    /// Returns where to scan for the trailer: past the entries of a classic <c>xref</c> table, whose entries are fixed
    /// 20-byte records, or at the cross-reference stream object itself.
    /// </summary>
    private static long LocatePdfTrailerDictionary(Stream stream, long start, long offset, long length,
        CancellationToken cancellationToken) {
        byte[] header = ReadPdfWindow(stream, start, offset, (int)Math.Min(PdfXrefSubsectionHeaderBytes, length - offset));
        int index = SkipPdfWhiteSpace(header, 0);
        if (!StartsWithPdfKeyword(header, index, "xref")) return offset;

        long position = offset + index + 4;
        for (int subsection = 0; subsection < MaximumPdfXrefSubsections && position < length; subsection++) {
            cancellationToken.ThrowIfCancellationRequested();
            header = ReadPdfWindow(stream, start, position, (int)Math.Min(PdfXrefSubsectionHeaderBytes, length - position));
            index = SkipPdfWhiteSpace(header, 0);
            if (StartsWithPdfKeyword(header, index, "trailer")) return position + index;
            if (!TryReadPdfInteger(header, ref index, out _)) return -1;
            index = SkipPdfWhiteSpace(header, index);
            if (!TryReadPdfInteger(header, ref index, out long count) || count < 0) return -1;
            while (index < header.Length && (header[index] == (byte)' ' || header[index] == (byte)'\t')) index++;
            if (index < header.Length && header[index] == (byte)'\r') index++;
            if (index < header.Length && header[index] == (byte)'\n') index++;
            // Entries are 20 bytes each; a count that cannot fit in the rest of the file is malformed.
            if (count > (length - position - index) / 20) return -1;
            position += index + count * 20;
        }
        return -1;
    }

    /// <summary>
    /// Scans PDF tokens once, skipping comments, strings and stream data, and reports an <c>Encrypt</c> key in a top-level
    /// dictionary that follows the <c>trailer</c> keyword or has <c>/Type /XRef</c>. Name escapes such as
    /// <c>/En#63rypt</c> are decoded before comparison.
    /// </summary>
    private static bool ContainsPdfTrailerEncryptEntry(byte[] bytes, CancellationToken cancellationToken) {
        int depth = 0;
        int arrayDepth = 0;
        bool afterTrailerKeyword = false;
        bool dictionaryIsTrailer = false;
        bool dictionaryIsCrossReferenceStream = false;
        bool dictionaryHasEncrypt = false;
        bool encryptValuePending = false;
        bool expectingValue = false;
        string? currentKey = null;
        int index = 0;
        while (index < bytes.Length) {
            if ((index & 0xFFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            byte current = bytes[index];
            if (IsPdfWhiteSpace(current)) {
                index++;
                continue;
            }
            switch (current) {
                case (byte)'%':
                    while (index < bytes.Length && bytes[index] != (byte)'\r' && bytes[index] != (byte)'\n') index++;
                    continue;
                case (byte)'(':
                    index = SkipPdfLiteralString(bytes, index);
                    RecordPdfValue(isNull: false);
                    CompletePdfValue();
                    continue;
                case (byte)'<' when index + 1 < bytes.Length && bytes[index + 1] == (byte)'<':
                    RecordPdfValue(isNull: false);
                    depth++;
                    index += 2;
                    if (depth == 1) {
                        dictionaryIsTrailer = afterTrailerKeyword;
                        dictionaryIsCrossReferenceStream = false;
                        dictionaryHasEncrypt = false;
                        encryptValuePending = false;
                        expectingValue = false;
                        currentKey = null;
                        arrayDepth = 0;
                    }
                    afterTrailerKeyword = false;
                    continue;
                case (byte)'<':
                    while (index < bytes.Length && bytes[index] != (byte)'>') index++;
                    index++;
                    RecordPdfValue(isNull: false);
                    CompletePdfValue();
                    continue;
                case (byte)'>' when index + 1 < bytes.Length && bytes[index + 1] == (byte)'>':
                    index += 2;
                    if (depth == 0) continue;
                    depth--;
                    if (depth == 0) {
                        if (dictionaryHasEncrypt && (dictionaryIsTrailer || dictionaryIsCrossReferenceStream)) return true;
                    } else {
                        CompletePdfValue();
                    }
                    continue;
                case (byte)'[':
                    RecordPdfValue(isNull: false);
                    if (depth > 0) arrayDepth++;
                    index++;
                    continue;
                case (byte)']':
                    index++;
                    if (arrayDepth > 0 && --arrayDepth == 0) CompletePdfValue();
                    continue;
                case (byte)'/': {
                    string name = ReadPdfName(bytes, ref index);
                    if (depth == 1 && arrayDepth == 0) {
                        if (!expectingValue) {
                            currentKey = name;
                            expectingValue = true;
                            encryptValuePending = name == "Encrypt";
                        } else {
                            RecordPdfValue(isNull: false);
                            if (currentKey == "Type" && name == "XRef") dictionaryIsCrossReferenceStream = true;
                            expectingValue = false;
                        }
                    }
                    afterTrailerKeyword = false;
                    continue;
                }
                case (byte)')' or (byte)'>' or (byte)'{' or (byte)'}':
                    index++;
                    continue;
            }

            int tokenStart = index;
            while (index < bytes.Length && IsPdfRegularCharacter(bytes[index])) index++;
            if (depth == 0) {
                afterTrailerKeyword = IsPdfToken(bytes, tokenStart, index, "trailer");
                if (IsPdfToken(bytes, tokenStart, index, "stream")) {
                    // Stream data is opaque; resume after its end marker or stop if it runs past the window.
                    int end = FindPdfBytes(bytes, "endstream", index);
                    if (end < 0) return false;
                    index = end + "endstream".Length;
                }
            } else if (depth == 1 && arrayDepth == 0) {
                RecordPdfValue(IsPdfToken(bytes, tokenStart, index, "null"));
                CompletePdfValue();
            }
        }
        return false;

        // A null value means the key is absent; any other value, such as an indirect reference, marks encryption.
        void RecordPdfValue(bool isNull) {
            if (depth != 1 || arrayDepth != 0 || !encryptValuePending) return;
            if (!isNull) dictionaryHasEncrypt = true;
            encryptValuePending = false;
        }

        // A number, string, reference part or nested dictionary ends the current value and makes the next name a key.
        void CompletePdfValue() {
            if (depth == 1 && arrayDepth == 0) expectingValue = false;
            if (depth == 0) afterTrailerKeyword = false;
        }
    }

    private static string ReadPdfName(byte[] bytes, ref int index) {
        var name = new StringBuilder();
        index++;
        while (index < bytes.Length && IsPdfRegularCharacter(bytes[index])) {
            if (bytes[index] == (byte)'#' && index + 2 < bytes.Length &&
                TryGetHexValue(bytes[index + 1], out int high) && TryGetHexValue(bytes[index + 2], out int low)) {
                name.Append((char)(high * 16 + low));
                index += 3;
            } else {
                name.Append((char)bytes[index]);
                index++;
            }
        }
        return name.ToString();
    }

    private static int SkipPdfLiteralString(byte[] bytes, int index) {
        int nesting = 0;
        while (index < bytes.Length) {
            byte current = bytes[index++];
            if (current == (byte)'\\') {
                index++;
            } else if (current == (byte)'(') {
                nesting++;
            } else if (current == (byte)')' && --nesting == 0) {
                break;
            }
        }
        return index;
    }

    private static bool TryGetHexValue(byte value, out int digit) {
        digit = value switch {
            >= (byte)'0' and <= (byte)'9' => value - '0',
            >= (byte)'a' and <= (byte)'f' => value - 'a' + 10,
            >= (byte)'A' and <= (byte)'F' => value - 'A' + 10,
            _ => -1
        };
        return digit >= 0;
    }

    private static bool IsPdfToken(byte[] bytes, int start, int end, string token) {
        if (end - start != token.Length) return false;
        for (int index = 0; index < token.Length; index++) {
            if (bytes[start + index] != token[index]) return false;
        }
        return true;
    }

    private static bool StartsWithPdfKeyword(byte[] bytes, int index, string keyword) =>
        index + keyword.Length <= bytes.Length && IsPdfToken(bytes, index, index + keyword.Length, keyword) &&
        (index + keyword.Length == bytes.Length || !IsPdfRegularCharacter(bytes[index + keyword.Length]));

    private static int SkipPdfWhiteSpace(byte[] bytes, int index) {
        while (index < bytes.Length && IsPdfWhiteSpace(bytes[index])) index++;
        return index;
    }

    private static bool TryReadPdfInteger(byte[] bytes, ref int index, out long value) {
        value = 0;
        int digits = 0;
        while (index < bytes.Length && bytes[index] >= (byte)'0' && bytes[index] <= (byte)'9' && digits < 18) {
            value = value * 10 + (bytes[index] - (byte)'0');
            index++;
            digits++;
        }
        return digits > 0;
    }

    private static int FindPdfBytes(byte[] bytes, string value, int from) {
        int last = bytes.Length - value.Length;
        for (int index = Math.Max(0, from); index <= last; index++) {
            int matched = 0;
            while (matched < value.Length && bytes[index + matched] == value[matched]) matched++;
            if (matched == value.Length) return index;
        }
        return -1;
    }

    private static byte[] ReadPdfWindow(Stream stream, long start, long offset, int count) {
        var window = new byte[Math.Max(0, count)];
        stream.Position = start + offset;
        int total = 0;
        while (total < window.Length) {
            int read = stream.Read(window, total, window.Length - total);
            if (read <= 0) break;
            total += read;
        }
        if (total == window.Length) return window;
        var exact = new byte[total];
        Array.Copy(window, exact, total);
        return exact;
    }

    private static bool TryReadPdfStartXref(byte[] tail, out long offset) {
        offset = -1;
        int keyword = -1;
        for (int index = tail.Length - "startxref".Length; index >= 0; index--) {
            if (IsPdfToken(tail, index, index + "startxref".Length, "startxref")) {
                keyword = index;
                break;
            }
        }
        if (keyword < 0) return false;
        int position = SkipPdfWhiteSpace(tail, keyword + "startxref".Length);
        return TryReadPdfInteger(tail, ref position, out offset);
    }

    private static bool IsPdfWhiteSpace(byte value) =>
        value is 0 or 9 or 10 or 12 or 13 or 32;

    private static bool IsPdfRegularCharacter(byte value) =>
        !IsPdfWhiteSpace(value) && value is not ((byte)'(' or (byte)')' or (byte)'<' or (byte)'>' or (byte)'[' or (byte)']'
            or (byte)'{' or (byte)'}' or (byte)'/' or (byte)'%');
}
