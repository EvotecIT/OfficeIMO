namespace OfficeIMO.Excel {
    internal static class ExcelDocumentLoadRouting {
        private static readonly byte[] ZipSignature = { 0x50, 0x4B };
        private static readonly byte[] OleCompoundSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        internal static bool IsLegacyXls(byte[] bytes, string? filePath) {
            if (HasOleCompoundSignature(bytes)) {
                return true;
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

        private static bool HasOleCompoundSignature(byte[] bytes) {
            if (bytes.Length < OleCompoundSignature.Length) {
                return false;
            }

            for (int i = 0; i < OleCompoundSignature.Length; i++) {
                if (bytes[i] != OleCompoundSignature[i]) {
                    return false;
                }
            }

            return true;
        }
    }
}
