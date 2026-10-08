using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegWriter {
    private static void LoadBlockLuma(
        byte[] rgba,
        int stride,
        int rowOffset,
        int rowStride,
        int width,
        int height,
        int bx,
        int by,
        int[] yBlock) {
        var i = 0;
        for (var y = 0; y < 8; y++) {
            var py = by + y;
            if (py >= height) py = height - 1;
            var row = py * rowStride + rowOffset;
            for (var x = 0; x < 8; x++) {
                var px = bx + x;
                if (px >= width) px = width - 1;
                var p = row + px * 4;
                var r = rgba[p + 0];
                var g = rgba[p + 1];
                var b = rgba[p + 2];
                var a = rgba[p + 3];
                if (a != 255) {
                    var inv = 255 - a;
                    r = (byte)((r * a + 255 * inv + 127) / 255);
                    g = (byte)((g * a + 255 * inv + 127) / 255);
                    b = (byte)((b * a + 255 * inv + 127) / 255);
                }

                var yv = (77 * r + 150 * g + 29 * b + 128) >> 8;
                yBlock[i] = yv - 128;
                i++;
            }
        }
    }

    private static void LoadBlockChroma(
        byte[] rgba,
        int stride,
        int rowOffset,
        int rowStride,
        int width,
        int height,
        int bx,
        int by,
        int sampleW,
        int sampleH,
        int[] cbBlock,
        int[] crBlock) {
        var i = 0;
        var count = sampleW * sampleH;
        for (var y = 0; y < 8; y++) {
            var baseY = by + y * sampleH;
            for (var x = 0; x < 8; x++) {
                var baseX = bx + x * sampleW;
                var sumR = 0;
                var sumG = 0;
                var sumB = 0;

                for (var sy = 0; sy < sampleH; sy++) {
                    var py = baseY + sy;
                    if (py >= height) py = height - 1;
                    var row = py * rowStride + rowOffset;
                    for (var sx = 0; sx < sampleW; sx++) {
                        var px = baseX + sx;
                        if (px >= width) px = width - 1;
                        var p = row + px * 4;
                        var r = rgba[p + 0];
                        var g = rgba[p + 1];
                        var b = rgba[p + 2];
                        var a = rgba[p + 3];
                        if (a != 255) {
                            var inv = 255 - a;
                            r = (byte)((r * a + 255 * inv + 127) / 255);
                            g = (byte)((g * a + 255 * inv + 127) / 255);
                            b = (byte)((b * a + 255 * inv + 127) / 255);
                        }
                        sumR += r;
                        sumG += g;
                        sumB += b;
                    }
                }

                var rAvg = (sumR + count / 2) / count;
                var gAvg = (sumG + count / 2) / count;
                var bAvg = (sumB + count / 2) / count;

                var cb = ((-43 * rAvg - 85 * gAvg + 128 * bAvg + 128) >> 8) + 128;
                var cr = ((128 * rAvg - 107 * gAvg - 21 * bAvg + 128) >> 8) + 128;

                cbBlock[i] = cb - 128;
                crBlock[i] = cr - 128;
                i++;
            }
        }
    }

    private static bool IsGrayscale(byte[] rgba, int width, int height, int stride, int rowOffset, int rowStride, CancellationToken cancellationToken) {
        for (var y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            var row = y * rowStride + rowOffset;
            for (var x = 0; x < width; x++) {
                var p = row + x * 4;
                var r = rgba[p + 0];
                var g = rgba[p + 1];
                var b = rgba[p + 2];
                var a = rgba[p + 3];
                if (a != 255) {
                    var inv = 255 - a;
                    r = (byte)((r * a + 255 * inv + 127) / 255);
                    g = (byte)((g * a + 255 * inv + 127) / 255);
                    b = (byte)((b * a + 255 * inv + 127) / 255);
                }
                if (r != g || r != b) return false;
            }
        }
        return true;
    }

}
