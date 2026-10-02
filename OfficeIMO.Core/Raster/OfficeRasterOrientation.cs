using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Shared exact pixel permutation for EXIF orientations and bounded WebP metadata reading.</summary>
internal static class OfficeRasterOrientation {
    internal static OfficeImageOrientation ReadWebpOrientation(byte[] bytes, CancellationToken cancellationToken) {
        if (bytes.Length < 12 || bytes[0] != 'R' || bytes[1] != 'I' || bytes[2] != 'F' || bytes[3] != 'F'
            || bytes[8] != 'W' || bytes[9] != 'E' || bytes[10] != 'B' || bytes[11] != 'P') return OfficeImageOrientation.Normal;
        int offset = 12;
        while (offset <= bytes.Length - 8) {
            cancellationToken.ThrowIfCancellationRequested();
            uint length = (uint)bytes[offset + 4] | ((uint)bytes[offset + 5] << 8)
                | ((uint)bytes[offset + 6] << 16) | ((uint)bytes[offset + 7] << 24);
            if (length > bytes.Length - offset - 8) return OfficeImageOrientation.Normal;
            if (bytes[offset] == 'E' && bytes[offset + 1] == 'X' && bytes[offset + 2] == 'I' && bytes[offset + 3] == 'F'
                && OfficeImageOrientationNormalizer.TryReadExifOrientationPayload(bytes, offset + 8, (int)length, out var orientation)) return orientation;
            long next = (long)offset + 8 + length + (length & 1);
            if (next > bytes.Length) break;
            offset = (int)next;
        }
        return OfficeImageOrientation.Normal;
    }

    internal static byte[] Apply(
        byte[] rgba,
        ref int width,
        ref int height,
        int orientation,
        CancellationToken cancellationToken, string dimensionsLimitMessage) {
        if (orientation <= 1) return rgba;
        var srcWidth = width;
        var srcHeight = height;
        var destWidth = (orientation >= 5 && orientation <= 8) ? srcHeight : srcWidth;
        var destHeight = (orientation >= 5 && orientation <= 8) ? srcWidth : srcHeight;
        var result = OfficeRasterGuards.AllocateRgba32(destWidth, destHeight, dimensionsLimitMessage);

        for (var y = 0; y < destHeight; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (var x = 0; x < destWidth; x++) {
                int sx;
                int sy;
                switch (orientation) {
                    case 2:
                        sx = srcWidth - 1 - x;
                        sy = y;
                        break;
                    case 3:
                        sx = srcWidth - 1 - x;
                        sy = srcHeight - 1 - y;
                        break;
                    case 4:
                        sx = x;
                        sy = srcHeight - 1 - y;
                        break;
                    case 5:
                        sx = y;
                        sy = x;
                        break;
                    case 6:
                        sx = y;
                        sy = srcHeight - 1 - x;
                        break;
                    case 7:
                        sx = srcWidth - 1 - y;
                        sy = srcHeight - 1 - x;
                        break;
                    case 8:
                        sx = srcWidth - 1 - y;
                        sy = x;
                        break;
                    default:
                        sx = x;
                        sy = y;
                        break;
                }

                var srcIndex = (sy * srcWidth + sx) * 4;
                var dstIndex = (y * destWidth + x) * 4;
                result[dstIndex + 0] = rgba[srcIndex + 0];
                result[dstIndex + 1] = rgba[srcIndex + 1];
                result[dstIndex + 2] = rgba[srcIndex + 2];
                result[dstIndex + 3] = rgba[srcIndex + 3];
            }
        }

        width = destWidth;
        height = destHeight;
        return result;
    }

}
