using OfficeIMO.Shared;

namespace OfficeIMO.Excel {
    internal static class ExcelDocumentLoadRouting {
        private static readonly byte[] ZipSignature = { 0x50, 0x4B };

        internal static bool IsLegacyXls(byte[] bytes, string? filePath) {
            if (OfficeCompoundDocumentDetector.HasCompoundSignature(bytes)) {
                return RouteCompoundDocument(bytes);
            }

            if (HasZipSignature(bytes)) {
                return false;
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

        private static bool RouteCompoundDocument(byte[] bytes) {
            OfficeCompoundDocumentDetector.DocumentKind kind = OfficeCompoundDocumentDetector.Detect(bytes, out _);
            switch (kind) {
                case OfficeCompoundDocumentDetector.DocumentKind.ExcelWorkbook:
                    return true;
                case OfficeCompoundDocumentDetector.DocumentKind.WordDocument:
                    throw new InvalidDataException("The input contains a legacy Word document, not an Excel workbook. Load it with OfficeIMO.Word.WordDocument.");
                case OfficeCompoundDocumentDetector.DocumentKind.EncryptedOpenXmlPackage:
                    throw new InvalidDataException("The input is a password-encrypted Office Open XML package. Use ExcelDocument.LoadEncrypted and provide its password.");
                case OfficeCompoundDocumentDetector.DocumentKind.Ambiguous:
                    throw new InvalidDataException("The OLE compound file contains more than one root Office document stream and cannot be routed safely.");
                default:
                    // Let the legacy reader produce its stable compound/signature diagnostics.
                    return true;
            }
        }
    }
}
