namespace OfficeIMO.Word {
    internal static class WordDocumentLoadRouting {
        private static readonly byte[] ZipSignature = { 0x50, 0x4B };
        private static readonly byte[] OleCompoundSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        internal const int SignatureLength = 8;

        internal static bool IsLegacyDoc(byte[] bytes, string? filePath) {
            if (HasZipSignature(bytes)) {
                return false;
            }

            if (HasOleCompoundSignature(bytes)) {
                return true;
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

        internal static bool HasLegacyDocSignature(byte[] bytes) {
            if (HasZipSignature(bytes)) {
                return false;
            }

            return HasOleCompoundSignature(bytes);
        }
    }
}
