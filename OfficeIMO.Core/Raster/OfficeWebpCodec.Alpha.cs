using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    /// <summary>
    /// Restores the separate VP8 alpha plane in place, including raw/lossless storage and all four filters.
    /// Compressed alpha uses the lossless stream's green channel with implicit canvas dimensions.
    /// </summary>
    private static bool TryApplyVp8Alpha(byte[] bytes, int offset, int length, int width, int height,
        byte[] rgba, long retainedManagedBytes, CancellationToken cancellationToken) {
        if (!OfficeImageReader.HasValidWebpAlphaHeader(bytes, offset, length) ||
            !OfficeRasterGuards.TryEnsurePixelCount(width, height, out int pixelCount) ||
            rgba.LongLength != pixelCount * 4L || retainedManagedBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
        int control = bytes[offset];
        int compression = control & 3;
        uint[]? argb = null;
        if (compression == 0) {
            if (length != pixelCount + 1) return false;
        } else {
            var budget = new Vp8lAllocationBudget();
            if (!budget.TryReserveBytes(retainedManagedBytes) ||
                !TryDecodeVp8lArgb(new LsbBitReader(bytes, offset + 1, length - 1), width, height,
                    budget, cancellationToken, out argb)) return false;
        }

        int filter = (control >> 2) & 3;
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                if ((x & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                int pixel = y * width + x;
                int channel = pixel * 4 + 3;
                int predictor = 0;
                if (filter != 0) {
                    if (y == 0) {
                        if (x != 0) predictor = rgba[channel - 4];
                    } else if (x == 0) {
                        predictor = rgba[channel - width * 4];
                    } else if (filter == 1) {
                        predictor = rgba[channel - 4];
                    } else if (filter == 2) {
                        predictor = rgba[channel - width * 4];
                    } else {
                        predictor = Math.Max(0, Math.Min(255, rgba[channel - 4] +
                            rgba[channel - width * 4] - rgba[channel - width * 4 - 4]));
                    }
                }
                int value = argb == null ? bytes[offset + 1 + pixel] : (byte)(argb[pixel] >> 8);
                rgba[channel] = unchecked((byte)(value + predictor));
            }
        }
        return true;
    }
}
