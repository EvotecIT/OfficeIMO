using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private static bool TryReadPng(byte[] data, CancellationToken cancellationToken, out OfficeImageInfo info) =>
        TryReadPng(data, cancellationToken, out info, out _);

    private static bool TryReadPng(
        byte[] data,
        CancellationToken cancellationToken,
        out OfficeImageInfo info,
        out OfficePngContainerValidation pngValidation) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        pngValidation = default;
        byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
        if (data.Length < 33 ||
            !StartsWith(data, signature) ||
            ReadInt32BigEndian(data, 8) != 13 ||
            GetAscii(data, 12, 4) != "IHDR" ||
            !HasValidPngIhdrFields(data)) {
            return false;
        }

        int width = ReadInt32BigEndian(data, 16);
        int height = ReadInt32BigEndian(data, 20);
        if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out _) ||
            !OfficePngContainerValidation.TryCreate(data, cancellationToken, out pngValidation)) {
            return false;
        }
        double dpiX = 96.0;
        double dpiY = 96.0;

        int offset = 8;
        while (offset + 12 <= data.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = ReadInt32BigEndian(data, offset);
            long chunkEnd = (long)offset + 12L + length;
            if (length < 0 || chunkEnd > data.Length) {
                break;
            }

            string type = GetAscii(data, offset + 4, 4);
            if (type == "pHYs" && length >= 9) {
                uint xPpm = ReadUInt32BigEndian(data, offset + 8);
                uint yPpm = ReadUInt32BigEndian(data, offset + 12);
                byte unit = data[offset + 16];
                if (unit == 1 && xPpm > 0 && yPpm > 0) {
                    dpiX = xPpm * 0.0254;
                    dpiY = yPpm * 0.0254;
                }

                break;
            }

            offset = (int)chunkEnd;
        }

        info = new OfficeImageInfo(OfficeImageFormat.Png, width, height, dpiX, dpiY);
        return width > 0 && height > 0;
    }

    private static bool HasValidPngIhdrFields(byte[] data) {
        byte bitDepth = data[24];
        byte colorType = data[25];
        bool validBitDepth = colorType switch {
            0 => bitDepth is 1 or 2 or 4 or 8 or 16,
            2 or 4 or 6 => bitDepth is 8 or 16,
            3 => bitDepth is 1 or 2 or 4 or 8,
            _ => false
        };
        return validBitDepth && data[26] == 0 && data[27] == 0 && data[28] <= 1;
    }

}
