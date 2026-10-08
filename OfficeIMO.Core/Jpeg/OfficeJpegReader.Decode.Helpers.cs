using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private const string JpegDimensionsLimitMessage = "JPEG dimensions exceed limits.";
    private static readonly int[] CrToR = new int[256];
    private static readonly int[] CrToG = new int[256];
    private static readonly int[] CbToG = new int[256];
    private static readonly int[] CbToB = new int[256];

    static OfficeJpegReader() {
        for (var i = 0; i < 256; i++) {
            var d = i - 128;
            CrToR[i] = (91881 * d + 32768) >> 16;
            // Round the complete green channel once, after adding both chroma terms.
            CrToG[i] = -46802 * d;
            CbToG[i] = -22554 * d + 32768;
            CbToB[i] = (116130 * d + 32768) >> 16;
        }
    }

    private static void WriteBlock(byte[] buffer, int stride, int blockX, int blockY, byte[] pixels) {
        var baseX = blockX * 8;
        var baseY = blockY * 8;
        for (var y = 0; y < 8; y++) {
            var row = (baseY + y) * stride + baseX;
            var src = y * 8;
            Buffer.BlockCopy(pixels, src, buffer, row, 8);
        }
    }

    private static byte ClampToByte(int value) {
        if (value <= 0) return 0;
        if (value >= 255) return 255;
        return (byte)value;
    }

    private static int FindScanEnd(OfficeByteView data, int start, CancellationToken cancellationToken) {
        // Stuffed bytes and restart markers advance by two, so offsets can skip checkpoint boundaries.
        var i = start;
        int scanSteps = 0;
        while (i + 1 < data.Length) {
            if ((scanSteps++ & 0x3FFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (data[i] == 0xFF) {
                var j = i + 1;
                SkipFillBytes(data, ref j, cancellationToken);
                if (j >= data.Length) return data.Length;
                var marker = data[j];
                if (marker == 0x00) {
                    i = j + 1;
                    continue;
                }
                if (marker >= 0xD0 && marker <= 0xD7) {
                    i = j + 1;
                    continue;
                }
                return i;
            }
            i++;
        }
        return data.Length;
    }

    private static bool TryReadAdobeTransform(OfficeByteView data, out int transform) {
        transform = 0;
        if (data.Length < 12) return false;
        if (data[0] != (byte)'A' || data[1] != (byte)'d' || data[2] != (byte)'o' || data[3] != (byte)'b' || data[4] != (byte)'e') {
            return false;
        }
        transform = data[11];
        return true;
    }

    private static byte[] ApplyOrientation(byte[] rgba, ref int width, ref int height,
        int orientation, CancellationToken cancellationToken) =>
        OfficeRasterOrientation.Apply(rgba, ref width, ref height, orientation, cancellationToken, JpegDimensionsLimitMessage);

    private static double[,] BuildCosTable() {
        var table = new double[8, 8];
        for (var x = 0; x < 8; x++) {
            for (var u = 0; u < 8; u++) {
                table[x, u] = Math.Cos(((2 * x + 1) * u * Math.PI) / 16.0);
            }
        }
        return table;
    }

    private static ushort ReadUInt16BE(OfficeByteView data, int offset) {
        return (ushort)((data[offset] << 8) | data[offset + 1]);
    }

    private static byte ApplyCmyk(int c, int k) {
        var v = c + k;
        if (v > 255) v = 255;
        return (byte)(255 - v);
    }

}
