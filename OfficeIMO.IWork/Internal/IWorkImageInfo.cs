namespace OfficeIMO.IWork.Internal;

internal static class IWorkImageInfo {
    internal static (int? Width, int? Height) Read(byte[] bytes, string mediaType) {
        if (string.Equals(mediaType, "image/png", StringComparison.OrdinalIgnoreCase)
            && bytes.Length >= 24
            && HasPngSignature(bytes)) {
            int width = ReadBigEndian32(bytes, 16);
            int height = ReadBigEndian32(bytes, 20);
            return width > 0 && height > 0 ? (width, height) : (null, null);
        }
        if (string.Equals(mediaType, "image/jpeg", StringComparison.OrdinalIgnoreCase)) return ReadJpeg(bytes);
        return (null, null);
    }

    private static (int? Width, int? Height) ReadJpeg(byte[] bytes) {
        if (bytes.Length < 4 || bytes[0] != 0xff || bytes[1] != 0xd8) return (null, null);
        int offset = 2;
        while (offset <= bytes.Length - 4) {
            if (bytes[offset++] != 0xff) continue;
            byte marker = bytes[offset++];
            while (marker == 0xff && offset < bytes.Length) marker = bytes[offset++];
            if (marker == 0xd9 || marker == 0xda) break;
            if (offset > bytes.Length - 2) break;
            int length = bytes[offset] << 8 | bytes[offset + 1];
            if (length < 2 || offset > bytes.Length - length) break;
            bool sizeMarker = marker >= 0xc0 && marker <= 0xc3
                || marker >= 0xc5 && marker <= 0xc7
                || marker >= 0xc9 && marker <= 0xcb
                || marker >= 0xcd && marker <= 0xcf;
            if (sizeMarker && length >= 7) {
                int height = bytes[offset + 3] << 8 | bytes[offset + 4];
                int width = bytes[offset + 5] << 8 | bytes[offset + 6];
                return width > 0 && height > 0 && HasJpegEndMarker(bytes, offset + length)
                    ? (width, height)
                    : (null, null);
            }
            offset += length;
        }
        return (null, null);
    }

    private static bool HasPngSignature(byte[] bytes) =>
        bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4e && bytes[3] == 0x47
        && bytes[4] == 0x0d && bytes[5] == 0x0a && bytes[6] == 0x1a && bytes[7] == 0x0a;

    private static bool HasJpegEndMarker(byte[] bytes, int offset) {
        for (int index = Math.Max(offset, 1); index < bytes.Length; index++) {
            if (bytes[index - 1] == 0xff && bytes[index] == 0xd9) return true;
        }
        return false;
    }

    private static int ReadBigEndian32(byte[] bytes, int offset) =>
        bytes[offset] << 24 | bytes[offset + 1] << 16 | bytes[offset + 2] << 8 | bytes[offset + 3];
}
