namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private static bool TryReadWebp(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 20 ||
            GetAscii(data, 0, 4) != "RIFF" ||
            GetAscii(data, 8, 4) != "WEBP") {
            return false;
        }

        long containerLength = 8L + ReadUInt32LittleEndian(data, 4);
        if (containerLength != data.LongLength) return false;

        uint chunkSize = ReadUInt32LittleEndian(data, 16);
        long chunkDataEnd = 20L + chunkSize;
        long paddedChunkEnd = chunkDataEnd + (chunkSize & 1U);
        if (chunkDataEnd > containerLength || paddedChunkEnd > containerLength) return false;

        string chunkType = GetAscii(data, 12, 4);
        int width;
        int height;
        if (chunkType == "VP8X" && chunkSize == 10U) {
            width = 1 + ReadUInt24LittleEndian(data, 24);
            height = 1 + ReadUInt24LittleEndian(data, 27);
        } else if (chunkType == "VP8L" && chunkSize >= 5U && data[20] == 0x2F) {
            width = 1 + data[21] + ((data[22] & 0x3F) << 8);
            height = 1 + ((data[22] & 0xC0) >> 6) + (data[23] << 2) + ((data[24] & 0x0F) << 10);
        } else if (chunkType == "VP8 " && chunkSize >= 10U &&
            data[23] == 0x9D && data[24] == 0x01 && data[25] == 0x2A) {
            width = ReadUInt16LittleEndian(data, 26) & 0x3FFF;
            height = ReadUInt16LittleEndian(data, 28) & 0x3FFF;
        } else {
            return false;
        }

        info = new OfficeImageInfo(OfficeImageFormat.Webp, width, height);
        return width > 0 && height > 0;
    }

    private static uint ReadUInt32LittleEndian(byte[] data, int offset) =>
        offset + 4 <= data.Length
            ? (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24))
            : 0U;
}
