namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private const ushort ZipEncryptedEntryFlag = 0x0001;
    private const int PdfStartXrefTailBytes = 2048;

    /// <summary>
    /// Adds encryption evidence that bounded reads can confirm without opening or decrypting the input: a PDF encryption
    /// dictionary in the active trailer and encrypted ZIP entries. An encrypted Office package takes its kind from a
    /// matching Word, Excel or PowerPoint extension at low confidence, because the format is only recorded inside the
    /// encrypted payload.
    /// </summary>
    private static DetectionCandidate ApplyContentEncryption(
        DetectionCandidate candidate,
        ReaderDetectionResult extensionResult,
        Stream stream,
        long start,
        bool seekable,
        byte[] prefix,
        ReaderDetectionOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (candidate.IsEncrypted) {
            return candidate.Kind == ReaderInputKind.Unknown &&
                   candidate.Evidence.Contains(EncryptedOpenXmlEvidence, StringComparer.Ordinal) &&
                   extensionResult.ExtensionKind is ReaderInputKind.Word or ReaderInputKind.Excel or ReaderInputKind.PowerPoint
                ? candidate.WithKind(extensionResult.ExtensionKind, ReaderDetectionConfidence.Low,
                    extensionResult.MediaType ?? GetMediaType(extensionResult.ExtensionKind))
                : candidate;
        }

        if (candidate.Kind == ReaderInputKind.Pdf) {
            return HasPdfEncryptDictionary(stream, start, seekable, prefix, options.MaxProbeBytes, cancellationToken)
                ? candidate.WithEncryption("encryption:pdf-encrypt-dictionary")
                : candidate;
        }

        if (candidate.Kind == ReaderInputKind.Zip) {
            bool encrypted = HasEncryptedFirstZipEntry(prefix) ||
                             seekable && options.InspectContainers &&
                             HasEncryptedZipEntry(stream, start, options.MaxContainerEntries, cancellationToken);
            return encrypted ? candidate.WithEncryption("encryption:zip-encrypted-entry") : candidate;
        }

        return candidate;
    }

    private static bool HasEncryptedFirstZipEntry(byte[] prefix) =>
        prefix.Length >= 8 && ReadUInt32(prefix, 0) == ZipLocalHeaderSignature &&
        (ReadUInt16(prefix, 6) & ZipEncryptedEntryFlag) != 0;

    private static bool HasEncryptedZipEntry(Stream stream, long start, int maxEntries, CancellationToken cancellationToken) {
        long originalPosition = stream.Position;
        try {
            if (!TryLocateZipCentralDirectory(stream, start, out long centralDirectoryOffset, out int entryCount)) return false;
            stream.Position = start + centralDirectoryOffset;
            var header = new byte[ZipCentralDirectoryHeaderLength];
            int entriesToInspect = Math.Min(entryCount, maxEntries);
            for (int entryIndex = 0; entryIndex < entriesToInspect; entryIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!ReadExact(stream, header, 0, header.Length) ||
                    ReadUInt32(header, 0) != ZipCentralDirectoryHeaderSignature) {
                    return false;
                }
                if ((ReadUInt16(header, 8) & ZipEncryptedEntryFlag) != 0) return true;
                long next = stream.Position + ReadUInt16(header, 28) + ReadUInt16(header, 30) + ReadUInt16(header, 32);
                if (next > stream.Length) return false;
                stream.Position = next;
            }
            return false;
        } finally {
            stream.Position = originalPosition;
        }
    }

    /// <summary>
    /// Looks for <c>/Encrypt</c> in the trailer dictionaries or cross-reference stream dictionaries visible in the prefix,
    /// the file tail and the cross-reference section named by the final <c>startxref</c>. Every trailer of an encrypted
    /// PDF carries the key, including the first-page trailer of a linearized file.
    /// </summary>
    private static bool HasPdfEncryptDictionary(Stream stream, long start, bool seekable, byte[] prefix, int windowBytes,
        CancellationToken cancellationToken) {
        if (ContainsPdfTrailerEncrypt(prefix)) return true;
        if (!seekable) return false;

        long originalPosition = stream.Position;
        try {
            long length = stream.Length - start;
            if (length <= prefix.Length) return false;
            byte[] tail = ReadPdfWindow(stream, start, Math.Max(0, length - windowBytes), windowBytes, length);
            if (ContainsPdfTrailerEncrypt(tail)) return true;
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryReadPdfStartXref(tail, out long crossReferenceOffset) || crossReferenceOffset >= length) return false;
            return ContainsPdfTrailerEncrypt(ReadPdfWindow(stream, start, crossReferenceOffset, windowBytes, length));
        } finally {
            stream.Position = originalPosition;
        }
    }

    private static byte[] ReadPdfWindow(Stream stream, long start, long offset, int windowBytes, long length) {
        int count = (int)Math.Min(windowBytes, length - offset);
        var window = new byte[count];
        stream.Position = start + offset;
        int total = 0;
        while (total < count) {
            int read = stream.Read(window, total, count - total);
            if (read <= 0) break;
            total += read;
        }
        if (total == count) return window;
        var exact = new byte[total];
        Array.Copy(window, exact, total);
        return exact;
    }

    private static bool TryReadPdfStartXref(byte[] tail, out long offset) {
        offset = -1;
        int searchFrom = Math.Max(0, tail.Length - PdfStartXrefTailBytes);
        int keyword = FindLastAscii(tail, "startxref", searchFrom);
        if (keyword < 0) return false;
        int index = keyword + "startxref".Length;
        while (index < tail.Length && IsPdfWhiteSpace(tail[index])) index++;
        long value = 0;
        int digits = 0;
        while (index < tail.Length && tail[index] >= (byte)'0' && tail[index] <= (byte)'9' && digits < 18) {
            value = value * 10 + (tail[index] - (byte)'0');
            index++;
            digits++;
        }
        if (digits == 0) return false;
        offset = value;
        return true;
    }

    private static bool ContainsPdfTrailerEncrypt(byte[] window) {
        // Classic trailers: the dictionary follows the "trailer" keyword.
        for (int index = FindAscii(window, "trailer", 0); index >= 0; index = FindAscii(window, "trailer", index + 7)) {
            int end = FindAscii(window, "startxref", index + 7);
            if (ContainsPdfEncryptKey(window, index + 7, end < 0 ? window.Length : end)) return true;
        }
        // Cross-reference streams: the trailer entries sit in the stream object's dictionary, ahead of its data.
        for (int index = FindAscii(window, "/XRef", 0); index >= 0; index = FindAscii(window, "/XRef", index + 5)) {
            int objectStart = FindLastAscii(window, " obj", 0, index);
            int streamStart = FindAscii(window, "stream", index);
            if (objectStart < 0) objectStart = Math.Max(0, index - 4096);
            if (ContainsPdfEncryptKey(window, objectStart, streamStart < 0 ? window.Length : streamStart)) return true;
        }
        return false;
    }

    private static bool ContainsPdfEncryptKey(byte[] window, int from, int to) {
        for (int index = FindAscii(window, "/Encrypt", from, to); index >= 0; index = FindAscii(window, "/Encrypt", index + 8, to)) {
            int next = index + 8;
            // "/EncryptMetadata" is a different key inside the encryption dictionary itself.
            if (next >= window.Length || !IsPdfRegularCharacter(window[next])) return true;
        }
        return false;
    }

    private static int FindAscii(byte[] source, string value, int from, int to = int.MaxValue) {
        int last = Math.Min(source.Length, to) - value.Length;
        for (int index = Math.Max(0, from); index <= last; index++) {
            int matched = 0;
            while (matched < value.Length && source[index + matched] == value[matched]) matched++;
            if (matched == value.Length) return index;
        }
        return -1;
    }

    private static int FindLastAscii(byte[] source, string value, int from, int to = int.MaxValue) {
        for (int index = Math.Min(source.Length, to) - value.Length; index >= Math.Max(0, from); index--) {
            int matched = 0;
            while (matched < value.Length && source[index + matched] == value[matched]) matched++;
            if (matched == value.Length) return index;
        }
        return -1;
    }

    private static bool IsPdfWhiteSpace(byte value) =>
        value is 0 or 9 or 10 or 12 or 13 or 32;

    private static bool IsPdfRegularCharacter(byte value) =>
        !IsPdfWhiteSpace(value) && value is not ((byte)'(' or (byte)')' or (byte)'<' or (byte)'>' or (byte)'[' or (byte)']'
            or (byte)'{' or (byte)'}' or (byte)'/' or (byte)'%');
}
