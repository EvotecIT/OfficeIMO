namespace OfficeIMO.OpenDocument;

internal static class OdfZipHeaderInspector {
    internal static void ValidateMimetypeEntry(byte[] packageBytes) {
        const uint localHeaderSignature = 0x04034b50;
        if (packageBytes.Length < 38 || ReadUInt32(packageBytes, 0) != localHeaderSignature) {
            throw new InvalidDataException("OpenDocument package does not begin with a valid ZIP local header.");
        }

        ushort compressionMethod = ReadUInt16(packageBytes, 8);
        ushort fileNameLength = ReadUInt16(packageBytes, 26);
        ushort extraFieldLength = ReadUInt16(packageBytes, 28);
        if (30 + fileNameLength + extraFieldLength > packageBytes.Length) {
            throw new InvalidDataException("OpenDocument package has a truncated first ZIP entry.");
        }

        string fileName = Encoding.UTF8.GetString(packageBytes, 30, fileNameLength);
        if (!string.Equals(fileName, "mimetype", StringComparison.Ordinal)) {
            throw new InvalidDataException("OpenDocument 'mimetype' must be the first ZIP entry.");
        }
        if (compressionMethod != 0) {
            throw new InvalidDataException("OpenDocument 'mimetype' must be stored without compression.");
        }
        if (extraFieldLength != 0) {
            throw new InvalidDataException("OpenDocument 'mimetype' must not use a ZIP extra field.");
        }
    }

    private static ushort ReadUInt16(byte[] bytes, int offset) => (ushort)(bytes[offset] | (bytes[offset + 1] << 8));

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        (uint)(bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24));
}
