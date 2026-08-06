namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private static bool TryReadJpeg(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 4 || data[0] != 0xFF || data[1] != 0xD8) {
            return false;
        }

        double dpiX = 96.0;
        double dpiY = 96.0;
        int offset = 2;

        while (offset < data.Length) {
            if (data[offset] != 0xFF) {
                return false;
            }

            while (offset < data.Length && data[offset] == 0xFF) {
                offset++;
            }

            if (offset >= data.Length) {
                break;
            }

            byte marker = data[offset++];
            if (marker == 0xD9 || marker == 0xDA) {
                return false;
            }
            if (marker == 0x01) {
                continue;
            }
            if (marker == 0x00 || marker == 0xD8 || (marker >= 0xD0 && marker <= 0xD7)) {
                return false;
            }

            if (offset + 2 > data.Length) {
                return false;
            }

            int segmentLength = ReadUInt16BigEndian(data, offset);
            if (segmentLength < 2 || offset + segmentLength > data.Length) {
                return false;
            }

            int segmentStart = offset + 2;
            int segmentDataLength = segmentLength - 2;

            if (marker == 0xE0 && segmentDataLength >= 12 && GetAscii(data, segmentStart, 5) == "JFIF\0") {
                byte units = data[segmentStart + 7];
                int xDensity = ReadUInt16BigEndian(data, segmentStart + 8);
                int yDensity = ReadUInt16BigEndian(data, segmentStart + 10);
                if (xDensity > 0 && yDensity > 0) {
                    if (units == 1) {
                        dpiX = xDensity;
                        dpiY = yDensity;
                    } else if (units == 2) {
                        dpiX = xDensity * 2.54;
                        dpiY = yDensity * 2.54;
                    }
                }
            }

            if (IsStartOfFrame(marker)) {
                if (!TryReadJpegFrameHeader(
                    data,
                    segmentStart,
                    segmentDataLength,
                    out int width,
                    out int height)) {
                    return false;
                }
                info = new OfficeImageInfo(OfficeImageFormat.Jpeg, width, height, dpiX, dpiY);
                return true;
            }

            offset += segmentLength;
        }

        return false;
    }

    private static bool TryReadJpegFrameHeader(
        byte[] data,
        int segmentStart,
        int segmentDataLength,
        out int width,
        out int height) {
        width = 0;
        height = 0;
        if (segmentDataLength < 9) return false;

        int componentCount = data[segmentStart + 5];
        if (componentCount == 0 || segmentDataLength != 6 + (componentCount * 3)) return false;
        if (data[segmentStart] == 0) return false;

        height = ReadUInt16BigEndian(data, segmentStart + 1);
        width = ReadUInt16BigEndian(data, segmentStart + 3);
        return width > 0 && height > 0;
    }

    private static bool IsStartOfFrame(byte marker) =>
        marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;
}
