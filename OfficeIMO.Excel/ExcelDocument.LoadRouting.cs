using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Excel {
    internal static class ExcelDocumentLoadRouting {
        private static readonly byte[] ZipSignature = { 0x50, 0x4B };

        internal static bool IsLegacyXls(byte[] bytes, string? filePath) {
            if (OfficeCompoundDocumentDetector.HasCompoundSignature(bytes)) {
                return RouteCompoundDocument(bytes, encryptedLoad: false, filePath);
            }

            if (HasZipSignature(bytes)) {
                return false;
            }

            return HasLegacyXlsExtension(filePath);
        }

        /// <summary>
        /// Distinguishes an encrypted legacy workbook from an encrypted Open XML package by its OLE root streams.
        /// </summary>
        internal static bool IsEncryptedLegacyXls(byte[] bytes, string? filePath) {
            if (OfficeCompoundDocumentDetector.HasCompoundSignature(bytes)) {
                return RouteCompoundDocument(bytes, encryptedLoad: true, filePath);
            }

            return HasLegacyXlsExtension(filePath);
        }

        internal static bool HasLegacyXlsExtension(string? filePath) {
            return !string.IsNullOrWhiteSpace(filePath)
                && string.Equals(Path.GetExtension(filePath), ".xls", StringComparison.OrdinalIgnoreCase);
        }

        internal static bool HasLegacyBinaryExcelExtension(string? filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                return false;
            }

            string extension = Path.GetExtension(filePath);
            return string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xlt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xla", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xlm", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xlw", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasZipSignature(byte[] bytes) {
            return bytes.Length >= ZipSignature.Length
                && bytes[0] == ZipSignature[0]
                && bytes[1] == ZipSignature[1];
        }

        private static bool RouteCompoundDocument(byte[] bytes, bool encryptedLoad, string? filePath) {
            OfficeCompoundDocumentDetector.DocumentKind kind = OfficeCompoundDocumentDetector.Detect(bytes, out _);
            switch (kind) {
                case OfficeCompoundDocumentDetector.DocumentKind.ExcelWorkbook:
                    return true;
                case OfficeCompoundDocumentDetector.DocumentKind.WordDocument:
                    throw new InvalidDataException("The input contains a legacy Word document, not an Excel workbook. Load it with OfficeIMO.Word.WordDocument.");
                case OfficeCompoundDocumentDetector.DocumentKind.EncryptedOpenXmlPackage:
                    if (encryptedLoad) {
                        return false;
                    }
                    throw new InvalidDataException("The input is a password-encrypted Office Open XML package. Use ExcelDocument.LoadEncrypted and provide its password.");
                case OfficeCompoundDocumentDetector.DocumentKind.Ambiguous:
                    throw new InvalidDataException("The OLE compound file contains more than one root Office document stream and cannot be routed safely.");
                default:
                    // Normal loads retain their legacy-reader diagnostics. Encrypted loads
                    // fall back to the extension only when the compound root is unknown.
                    return !encryptedLoad || HasLegacyXlsExtension(filePath);
            }
        }
    }
}
