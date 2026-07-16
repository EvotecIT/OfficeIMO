namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private const int PcxHeaderSizeBytes = 128;
    private const int PcxExtendedPaletteSizeBytes = 769;

    private static bool TryReadPcx(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < PcxHeaderSizeBytes ||
            data[0] != 0x0A ||
            data[1] is not (0 or 2 or 3 or 4 or 5) ||
            data[2] != 0x01 ||
            data[64] != 0) {
            return false;
        }

        int xMin = ReadUInt16LittleEndian(data, 4);
        int yMin = ReadUInt16LittleEndian(data, 6);
        int xMax = ReadUInt16LittleEndian(data, 8);
        int yMax = ReadUInt16LittleEndian(data, 10);
        if (xMax < xMin || yMax < yMin) {
            return false;
        }

        int width = xMax - xMin + 1;
        int height = yMax - yMin + 1;
        int bitsPerPixel = data[3];
        int planes = data[65];
        bool supportedLayout = bitsPerPixel switch {
            1 => planes is >= 1 and <= 4,
            2 or 4 => planes == 1,
            8 => planes is 1 or 3 or 4,
            _ => false
        };
        int bytesPerLine = ReadUInt16LittleEndian(data, 66);
        int minimumBytesPerLine = (width * bitsPerPixel + 7) / 8;
        if (!supportedLayout || bytesPerLine < minimumBytesPerLine || (bytesPerLine & 1) != 0 ||
            !HasCompletePcxImageData(data, height, planes, bytesPerLine, data[1], bitsPerPixel)) {
            return false;
        }

        double dpiX = ReadUInt16LittleEndian(data, 12);
        double dpiY = ReadUInt16LittleEndian(data, 14);

        info = new OfficeImageInfo(OfficeImageFormat.Pcx, width, height, dpiX, dpiY);
        return width > 0 && height > 0;
    }

    private static bool HasCompletePcxImageData(
        byte[] data,
        int height,
        int planes,
        int bytesPerLine,
        int version,
        int bitsPerPixel) {
        bool requiresExtendedPalette = version == 5 && bitsPerPixel == 8 && planes == 1;
        int encodedEnd = data.Length;
        if (requiresExtendedPalette) {
            if (data.Length < PcxHeaderSizeBytes + PcxExtendedPaletteSizeBytes) {
                return false;
            }

            encodedEnd -= PcxExtendedPaletteSizeBytes;
            if (data[encodedEnd] != 0x0C) {
                return false;
            }
        }

        int encodedOffset = PcxHeaderSizeBytes;
        int decodedBytesPerScanline = checked(planes * bytesPerLine);
        for (int row = 0; row < height; row++) {
            int remaining = decodedBytesPerScanline;
            while (remaining > 0) {
                if (encodedOffset >= encodedEnd) {
                    return false;
                }

                int value = data[encodedOffset++];
                int count = 1;
                if ((value & 0xC0) == 0xC0) {
                    count = value & 0x3F;
                    if (count == 0 || encodedOffset >= encodedEnd) {
                        return false;
                    }

                    encodedOffset++;
                }

                if (count > remaining) {
                    return false;
                }

                remaining -= count;
            }
        }

        return encodedOffset == encodedEnd;
    }
}
