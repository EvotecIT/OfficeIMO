using System;

namespace OfficeIMO.Drawing.Internal {
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
            UnknownCompound
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

            if (!OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compoundFile, out error) || compoundFile == null) {
                return DocumentKind.UnknownCompound;
            }

            bool hasWordDocument = compoundFile.Streams.ContainsKey("WordDocument");
            bool hasWorkbook = compoundFile.Streams.ContainsKey("Workbook") || compoundFile.Streams.ContainsKey("Book");
            bool hasPowerPointPresentation = compoundFile.Streams.ContainsKey("PowerPoint Document")
                && compoundFile.Streams.ContainsKey("Current User");
            bool hasEncryptedPackage = compoundFile.Streams.ContainsKey("EncryptedPackage")
                && compoundFile.Streams.ContainsKey("EncryptionInfo");

            int recognizedRootCount = (hasWordDocument ? 1 : 0)
                + (hasWorkbook ? 1 : 0)
                + (hasPowerPointPresentation ? 1 : 0)
                + (hasEncryptedPackage ? 1 : 0);
            if (recognizedRootCount > 1) return DocumentKind.Ambiguous;
            if (hasWordDocument) return DocumentKind.WordDocument;
            if (hasWorkbook) return DocumentKind.ExcelWorkbook;
            if (hasPowerPointPresentation) return DocumentKind.PowerPointPresentation;
            if (hasEncryptedPackage) return DocumentKind.EncryptedOpenXmlPackage;
            return DocumentKind.UnknownCompound;
        }
    }
}
