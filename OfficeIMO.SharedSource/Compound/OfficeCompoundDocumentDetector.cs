using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

#if OFFICEIMO_READER_CORE
namespace OfficeIMO.Reader.Internal.Compound {
#else
namespace OfficeIMO.Core.Internal {
#endif
    /// <summary>
    /// Identifies the root document payload in an OLE compound file.
    /// </summary>
    internal static class OfficeCompoundDocumentDetector {
        private static readonly byte[] OleCompoundSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        internal enum DocumentKind {
            NotCompound,
            WordDocument,
            ExcelWorkbook,
            PowerPointPresentation,
            EncryptedOpenXmlPackage,
            Ambiguous,
            UnknownCompound,
            PublisherPublication
        }

        internal static bool HasCompoundSignature(byte[] bytes) {
            if (bytes == null || bytes.Length < OleCompoundSignature.Length) return false;

            for (int i = 0; i < OleCompoundSignature.Length; i++) {
                if (bytes[i] != OleCompoundSignature[i]) return false;
            }

            return true;
        }

        internal static DocumentKind Detect(byte[] bytes, out string? error) {
            error = null;
            if (!HasCompoundSignature(bytes)) return DocumentKind.NotCompound;

            using var stream = new MemoryStream(bytes, writable: false);
            return Detect(stream, bytes.LongLength, 65536, out error);
        }

        /// <summary>
        /// Identifies the root Office document payload by inspecting only compound-file directory metadata.
        /// The source position is restored before this method returns.
        /// </summary>
        internal static DocumentKind Detect(Stream stream, long maxInputBytes,
            int maxDirectoryEntries, out string? error) {
            return Detect(stream, maxInputBytes, maxDirectoryEntries,
                CancellationToken.None, out error);
        }

        internal static DocumentKind Detect(Stream stream, long maxInputBytes,
            int maxDirectoryEntries, CancellationToken cancellationToken,
            out string? error) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            cancellationToken.ThrowIfCancellationRequested();
            if (!stream.CanRead || !stream.CanSeek) {
                error = "Compound document detection requires a readable seekable stream.";
                return DocumentKind.UnknownCompound;
            }

            long originalPosition = stream.Position;
            try {
                var signature = new byte[OleCompoundSignature.Length];
                int read = stream.Read(signature, 0, signature.Length);
                cancellationToken.ThrowIfCancellationRequested();
                if (read != signature.Length || !HasCompoundSignature(signature)) {
                    error = null;
                    return DocumentKind.NotCompound;
                }

                stream.Position = originalPosition;
                if (!OfficeCompoundFileReader.TryInspectDirectory(stream,
                        maxInputBytes, maxDirectoryEntries,
                        cancellationToken,
                        out IReadOnlyList<OfficeCompoundFileEntry> entries,
                        out error)) {
                    return DocumentKind.UnknownCompound;
                }

                return Detect(entries, cancellationToken);
            } finally {
                stream.Position = originalPosition;
            }
        }

        private static DocumentKind Detect(
            IReadOnlyList<OfficeCompoundFileEntry> entries,
            CancellationToken cancellationToken = default) {
            bool hasWordDocument = ContainsPayloadStream(entries, "WordDocument",
                cancellationToken);
            bool hasWorkbook = ContainsPayloadStream(entries, "Workbook",
                    cancellationToken)
                || ContainsPayloadStream(entries, "Book", cancellationToken);
            bool hasPowerPointPresentation =
                ContainsPayloadStream(entries, "PowerPoint Document",
                    cancellationToken)
                && ContainsPayloadStream(entries, "Current User",
                    cancellationToken);
            bool hasEncryptedPackage = ContainsPayloadStream(entries,
                    "EncryptedPackage", cancellationToken)
                && ContainsPayloadStream(entries, "EncryptionInfo",
                    cancellationToken);
            bool hasPublisher = ContainsPayloadStream(entries, "Contents", cancellationToken)
                && ContainsPayloadStream(entries, "Quill/QuillSub/CONTENTS", cancellationToken);

            int recognizedRootCount = (hasWordDocument ? 1 : 0)
                + (hasWorkbook ? 1 : 0)
                + (hasPowerPointPresentation ? 1 : 0)
                + (hasEncryptedPackage ? 1 : 0)
                + (hasPublisher ? 1 : 0);
            if (recognizedRootCount > 1) return DocumentKind.Ambiguous;
            if (hasWordDocument) return DocumentKind.WordDocument;
            if (hasWorkbook) return DocumentKind.ExcelWorkbook;
            if (hasPowerPointPresentation) return DocumentKind.PowerPointPresentation;
            if (hasEncryptedPackage) return DocumentKind.EncryptedOpenXmlPackage;
            if (hasPublisher) return DocumentKind.PublisherPublication;
            return DocumentKind.UnknownCompound;
        }

        private static bool ContainsPayloadStream(
            IEnumerable<OfficeCompoundFileEntry> entries, string name,
            CancellationToken cancellationToken) {
            foreach (OfficeCompoundFileEntry entry in entries) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!entry.IsFallback && entry.IsStream
                    && string.Equals(entry.Path, name,
                        StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }
            return false;
        }
    }
}
