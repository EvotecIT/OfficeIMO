using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointPresentationLoadRouting {
        internal static bool IsLegacyBinary(byte[] bytes, string? filePath) {
            if (OfficeCompoundDocumentDetector.HasCompoundSignature(bytes)) {
                OfficeCompoundDocumentDetector.DocumentKind kind = OfficeCompoundDocumentDetector.Detect(bytes, out _);
                switch (kind) {
                    case OfficeCompoundDocumentDetector.DocumentKind.PowerPointPresentation:
                        return true;
                    case OfficeCompoundDocumentDetector.DocumentKind.WordDocument:
                        throw new InvalidDataException("The input contains a legacy Word document, not a PowerPoint presentation. Load it with OfficeIMO.Word.WordDocument.");
                    case OfficeCompoundDocumentDetector.DocumentKind.ExcelWorkbook:
                        throw new InvalidDataException("The input contains a legacy Excel workbook, not a PowerPoint presentation. Load it with OfficeIMO.Excel.ExcelDocument.");
                    case OfficeCompoundDocumentDetector.DocumentKind.EncryptedOpenXmlPackage:
                        throw new InvalidDataException("The input is a password-encrypted Office Open XML package. Use PowerPointPresentation.LoadEncrypted and provide its password.");
                    case OfficeCompoundDocumentDetector.DocumentKind.Ambiguous:
                        throw new InvalidDataException("The OLE compound file contains more than one root Office document stream and cannot be routed safely.");
                    default:
                        return HasLegacyBinaryExtension(filePath);
                }
            }
            return HasLegacyBinaryExtension(filePath) && !HasZipSignature(bytes);
        }

        internal static bool HasLegacyBinaryExtension(string? filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) return false;
            string extension = Path.GetExtension(filePath);
            return string.Equals(extension, ".ppt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".pot", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".pps", StringComparison.OrdinalIgnoreCase);
        }

        internal static PowerPointFileFormat GetFormat(string? filePath, bool legacyDefault = false) {
            string extension = string.IsNullOrWhiteSpace(filePath) ? string.Empty : Path.GetExtension(filePath);
            if (string.Equals(extension, ".pot", StringComparison.OrdinalIgnoreCase)) return PowerPointFileFormat.Pot;
            if (string.Equals(extension, ".pps", StringComparison.OrdinalIgnoreCase)) return PowerPointFileFormat.Pps;
            if (string.Equals(extension, ".ppt", StringComparison.OrdinalIgnoreCase)) return PowerPointFileFormat.Ppt;
            return legacyDefault ? PowerPointFileFormat.Ppt : PowerPointFileFormat.Pptx;
        }

        private static bool HasZipSignature(byte[] bytes) => bytes.Length >= 2 && bytes[0] == 0x50 && bytes[1] == 0x4B;
    }
}
