namespace OfficeIMO.IWork.Internal;

internal static class IWorkImageInfo {
    internal static (int? Width, int? Height) Read(byte[] bytes, string mediaType) {
        if (string.Equals(mediaType, "image/png", StringComparison.OrdinalIgnoreCase)
            && TryReadPng(bytes, out int width, out int height)) {
            return (width, height);
        }
        if (string.Equals(mediaType, "image/jpeg", StringComparison.OrdinalIgnoreCase)) return ReadJpeg(bytes);
        return (null, null);
    }

    private static (int? Width, int? Height) ReadJpeg(byte[] bytes) {
        if (bytes.Length < 4 || bytes[0] != 0xff || bytes[1] != 0xd8) return (null, null);
        int offset = 2;
        int width = 0;
        int height = 0;
        bool hasScan = false;
        while (offset < bytes.Length) {
            if (bytes[offset++] != 0xff) return (null, null);
            if (offset >= bytes.Length) return (null, null);
            byte marker = bytes[offset++];
            while (marker == 0xff && offset < bytes.Length) marker = bytes[offset++];
            if (marker == 0xd9) return hasScan && width > 0 && height > 0
                ? (width, height)
                : (null, null);
            if (marker == 0x00 || marker == 0xd8) return (null, null);
            if (marker == 0x01 || marker >= 0xd0 && marker <= 0xd7) continue;
            if (offset > bytes.Length - 2) return (null, null);
            int length = bytes[offset] << 8 | bytes[offset + 1];
            if (length < 2 || offset > bytes.Length - length) return (null, null);
            bool sizeMarker = marker >= 0xc0 && marker <= 0xc3
                || marker >= 0xc5 && marker <= 0xc7
                || marker >= 0xc9 && marker <= 0xcb
                || marker >= 0xcd && marker <= 0xcf;
            if (sizeMarker && length >= 7) {
                height = bytes[offset + 3] << 8 | bytes[offset + 4];
                width = bytes[offset + 5] << 8 | bytes[offset + 6];
                if (width <= 0 || height <= 0) return (null, null);
            }
            offset += length;
            if (marker == 0xda) {
                hasScan = true;
                if (!SeekNextJpegMarker(bytes, ref offset)) return (null, null);
            }
        }
        return (null, null);
    }

    private static bool SeekNextJpegMarker(byte[] bytes, ref int offset) {
        while (offset < bytes.Length) {
            if (bytes[offset++] != 0xff) continue;
            int markerOffset = offset - 1;
            if (offset >= bytes.Length) return false;
            byte marker = bytes[offset++];
            while (marker == 0xff && offset < bytes.Length) marker = bytes[offset++];
            if (marker == 0x00 || marker >= 0xd0 && marker <= 0xd7) continue;
            offset = markerOffset;
            return true;
        }
        return false;
    }

    private static bool HasPngSignature(byte[] bytes) =>
        bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4e && bytes[3] == 0x47
        && bytes[4] == 0x0d && bytes[5] == 0x0a && bytes[6] == 0x1a && bytes[7] == 0x0a;

    private static bool TryReadPng(byte[] bytes, out int width, out int height) {
        width = 0;
        height = 0;
        if (bytes.Length < 33 || !HasPngSignature(bytes)) return false;

        bool hasHeader = false;
        bool hasImageData = false;
        int offset = 8;
        while (offset <= bytes.Length - 12) {
            uint rawLength = ReadBigEndianUInt32(bytes, offset);
            if (rawLength > int.MaxValue) return false;
            int dataLength = (int)rawLength;
            if (dataLength > bytes.Length - offset - 12) return false;

            int typeOffset = offset + 4;
            int dataOffset = typeOffset + 4;
            int crcOffset = dataOffset + dataLength;
            uint expectedCrc = ReadBigEndianUInt32(bytes, crcOffset);
            if (expectedCrc != CalculatePngCrc(bytes, typeOffset, checked(4 + dataLength))) return false;

            bool isHeader = IsChunk(bytes, typeOffset, 'I', 'H', 'D', 'R');
            bool isImageData = IsChunk(bytes, typeOffset, 'I', 'D', 'A', 'T');
            bool isEnd = IsChunk(bytes, typeOffset, 'I', 'E', 'N', 'D');
            if (!hasHeader) {
                if (!isHeader || offset != 8 || dataLength != 13) return false;
                width = ReadBigEndian32(bytes, dataOffset);
                height = ReadBigEndian32(bytes, dataOffset + 4);
                if (width <= 0 || height <= 0) return false;
                hasHeader = true;
            } else if (isHeader) {
                return false;
            }

            if (isImageData && dataLength > 0) hasImageData = true;
            offset = crcOffset + 4;
            if (isEnd) return dataLength == 0 && hasImageData && offset == bytes.Length;
        }
        return false;
    }

    private static bool IsChunk(byte[] bytes, int offset, char first, char second, char third, char fourth) =>
        bytes[offset] == first && bytes[offset + 1] == second
        && bytes[offset + 2] == third && bytes[offset + 3] == fourth;

    private static uint CalculatePngCrc(byte[] bytes, int offset, int length) {
        uint crc = uint.MaxValue;
        for (int index = offset; index < offset + length; index++) {
            crc ^= bytes[index];
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1) != 0 ? 0xedb88320U ^ crc >> 1 : crc >> 1;
            }
        }
        return crc ^ uint.MaxValue;
    }

    private static int ReadBigEndian32(byte[] bytes, int offset) =>
        bytes[offset] << 24 | bytes[offset + 1] << 16 | bytes[offset + 2] << 8 | bytes[offset + 3];

    private static uint ReadBigEndianUInt32(byte[] bytes, int offset) =>
        (uint)bytes[offset] << 24 | (uint)bytes[offset + 1] << 16
        | (uint)bytes[offset + 2] << 8 | bytes[offset + 3];
}
