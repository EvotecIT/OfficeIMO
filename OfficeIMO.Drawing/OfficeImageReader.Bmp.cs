using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private static bool TryReadBmp(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 26 || data[0] != (byte)'B' || data[1] != (byte)'M') {
            return false;
        }

        int dibSize = ReadInt32LittleEndian(data, 14);
        if (dibSize < 12) {
            return false;
        }

        int width;
        int height;
        double dpiX = 96.0;
        double dpiY = 96.0;

        if (dibSize == 12) {
            int planes = ReadUInt16LittleEndian(data, 22);
            int bitsPerPixel = ReadUInt16LittleEndian(data, 24);
            if (planes != 1 || bitsPerPixel is not (1 or 4 or 8 or 24)) {
                return false;
            }

            width = ReadUInt16LittleEndian(data, 18);
            height = ReadUInt16LittleEndian(data, 20);
        } else if (dibSize >= 40 && 14L + dibSize <= data.LongLength) {
            int planes = ReadUInt16LittleEndian(data, 26);
            int bitsPerPixel = ReadUInt16LittleEndian(data, 28);
            int compression = ReadInt32LittleEndian(data, 30);
            int rawHeight = ReadInt32LittleEndian(data, 22);
            if (planes != 1 || !HasSupportedBmpCompression(bitsPerPixel, compression, rawHeight)) {
                return false;
            }

            if (!TryConvertPixelDimension(ReadInt32LittleEndian(data, 18), out width) ||
                !TryConvertPixelDimension(Math.Abs((long)rawHeight), out height)) {
                return false;
            }
            int xPpm = ReadInt32LittleEndian(data, 38);
            int yPpm = ReadInt32LittleEndian(data, 42);
            if (xPpm > 0) dpiX = xPpm * 0.0254;
            if (yPpm > 0) dpiY = yPpm * 0.0254;
        } else {
            return false;
        }

        info = new OfficeImageInfo(OfficeImageFormat.Bmp, width, height, dpiX, dpiY);
        return width > 0 && height > 0;
    }

    private static bool HasSupportedBmpCompression(int bitsPerPixel, int compression, int rawHeight) {
        bool supported = bitsPerPixel switch {
            0 => compression is 4 or 5,
            1 => compression == 0,
            4 => compression is 0 or 2,
            8 => compression is 0 or 1,
            16 or 32 => compression is 0 or 3 or 6,
            24 => compression == 0,
            _ => false
        };

        if (!supported) {
            return false;
        }

        return rawHeight >= 0 || compression is 0 or 3;
    }
}
