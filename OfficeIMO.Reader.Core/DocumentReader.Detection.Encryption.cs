namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private const ushort ZipEncryptedEntryFlag = 0x0001;

    /// <summary>
    /// Adds encryption evidence that bounded reads can confirm without opening or decrypting the input: a PDF encryption
    /// dictionary in the active trailer, encrypted entries declared by an OpenDocument manifest, and encrypted ZIP entries.
    /// An encrypted Office package takes its kind from a matching Word, Excel or PowerPoint extension at low confidence,
    /// because the format is only recorded inside the encrypted payload.
    /// </summary>
    private static DetectionCandidate ApplyContentEncryption(
        DetectionCandidate candidate,
        ReaderDetectionResult extensionResult,
        Stream stream,
        long start,
        bool seekable,
        byte[] prefix,
        ReaderDetectionOptions options,
        bool zipContainer,
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
            return HasPdfEncryptDictionary(stream, start, seekable, prefix, cancellationToken)
                ? candidate.WithEncryption("encryption:pdf-encrypt-dictionary")
                : candidate;
        }

        if (candidate.Kind == ReaderInputKind.OpenDocument &&
            candidate.Evidence.Contains("container:opendocument-mimetype", StringComparer.Ordinal) &&
            seekable && options.InspectContainers &&
            HasEncryptedOpenDocumentEntries(stream, start, options.MaxContainerEntries, cancellationToken)) {
            return candidate.WithEncryption("encryption:opendocument-manifest");
        }

        // A ZIP-based package keeps its specialized kind, such as Word or EPUB, and still reports encrypted entries.
        if (zipContainer) {
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
}
