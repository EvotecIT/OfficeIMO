using OfficeIMO.Shared;

namespace OfficeIMO.Word {
    internal static class WordDocumentLoadRouting {
        private static readonly byte[] ZipSignature = { 0x50, 0x4B };
        internal const int SignatureLength = 8;

        internal static bool IsLegacyDoc(byte[] bytes, string? filePath) {
            if (HasZipSignature(bytes)) {
                return false;
            }

            if (OfficeCompoundDocumentDetector.HasCompoundSignature(bytes)) {
                return RouteCompoundDocument(bytes);
            }

            return HasLegacyDocExtension(filePath);
        }

        internal static bool HasLegacyDocExtension(string? filePath) {
            return !string.IsNullOrWhiteSpace(filePath)
                && string.Equals(Path.GetExtension(filePath), ".doc", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasZipSignature(byte[] bytes) {
            return bytes.Length >= ZipSignature.Length
                && bytes[0] == ZipSignature[0]
                && bytes[1] == ZipSignature[1];
        }

        internal static bool HasOleCompoundSignature(byte[] bytes) {
            return !HasZipSignature(bytes) && OfficeCompoundDocumentDetector.HasCompoundSignature(bytes);
        }

        private static bool RouteCompoundDocument(byte[] bytes) {
            OfficeCompoundDocumentDetector.DocumentKind kind = OfficeCompoundDocumentDetector.Detect(bytes, out string? error);
            switch (kind) {
                case OfficeCompoundDocumentDetector.DocumentKind.WordDocument:
                    return true;
                case OfficeCompoundDocumentDetector.DocumentKind.ExcelWorkbook:
                    throw new InvalidDataException("The input contains a legacy Excel workbook, not a Word document. Load it with OfficeIMO.Excel.ExcelDocument.");
                case OfficeCompoundDocumentDetector.DocumentKind.EncryptedOpenXmlPackage:
                    throw new InvalidDataException("The input is a password-encrypted Office Open XML package. Use WordDocument.LoadEncrypted and provide its password.");
                case OfficeCompoundDocumentDetector.DocumentKind.Ambiguous:
                    throw new InvalidDataException("The OLE compound file contains more than one root Office document stream and cannot be routed safely.");
                default:
                    throw new InvalidDataException(string.IsNullOrWhiteSpace(error)
                        ? "The OLE compound file does not contain a recognizable Word document stream."
                        : "The OLE compound file could not be identified safely. " + error);
            }
        }
    }
}
