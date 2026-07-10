namespace OfficeIMO.OpenDocument;

internal static class OdfZipHeaderInspector {
    internal static void ValidateMimetypeEntry(Stream packageStream) {
        if (packageStream == null) throw new ArgumentNullException(nameof(packageStream));
        if (!packageStream.CanRead || !packageStream.CanSeek) {
            throw new ArgumentException("OpenDocument ZIP header inspection requires a readable, seekable stream.", nameof(packageStream));
        }

        const uint localHeaderSignature = 0x04034b50;
        long originalPosition = packageStream.Position;
        try {
            var header = new byte[30];
            ReadExactly(packageStream, header, "OpenDocument package does not begin with a complete ZIP local header.");
            if (ReadUInt32(header, 0) != localHeaderSignature) {
                throw new InvalidDataException("OpenDocument package does not begin with a valid ZIP local header.");
            }

            ushort compressionMethod = ReadUInt16(header, 8);
            ushort fileNameLength = ReadUInt16(header, 26);
            ushort extraFieldLength = ReadUInt16(header, 28);
            if (fileNameLength == 0 || packageStream.Length - packageStream.Position < fileNameLength + extraFieldLength) {
                throw new InvalidDataException("OpenDocument package has a truncated first ZIP entry.");
            }

            var fileNameBytes = new byte[fileNameLength];
            ReadExactly(packageStream, fileNameBytes, "OpenDocument package has a truncated first ZIP entry.");
            string fileName = Encoding.UTF8.GetString(fileNameBytes);
            if (!string.Equals(fileName, "mimetype", StringComparison.Ordinal)) {
                throw new InvalidDataException("OpenDocument 'mimetype' must be the first ZIP entry.");
            }
            if (compressionMethod != 0) {
                throw new InvalidDataException("OpenDocument 'mimetype' must be stored without compression.");
            }
            if (extraFieldLength != 0) {
                throw new InvalidDataException("OpenDocument 'mimetype' must not use a ZIP extra field.");
            }
        } finally {
            packageStream.Position = originalPosition;
        }
    }

    private static void ReadExactly(Stream stream, byte[] buffer, string error) {
        int offset = 0;
        while (offset < buffer.Length) {
            int read = stream.Read(buffer, offset, buffer.Length - offset);
            if (read == 0) throw new InvalidDataException(error);
            offset += read;
        }
    }

    private static ushort ReadUInt16(byte[] bytes, int offset) => (ushort)(bytes[offset] | (bytes[offset + 1] << 8));

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        (uint)(bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24));
}
