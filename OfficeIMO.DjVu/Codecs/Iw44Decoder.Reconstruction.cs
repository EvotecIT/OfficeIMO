namespace OfficeIMO.DjVu;

internal sealed partial class Iw44Decoder {
    private int[] ReconstructPlane(Iw44Plane plane, int minimumScale) {
        var samples = new int[checked(Width * Height)];
        int blocksAcross = (Width + 31) / 32;
        for (int block = 0; block < plane.Coefficients.Length / 1024; block++) {
            _budget.Cancellation.ThrowIfCancellationRequested();
            int bx = block % blocksAcross * 32, by = block / blocksAcross * 32;
            for (int i = 0; i < 1024; i++) {
                int x = 0, y = 0;
                for (int bit = 0; bit < 5; bit++) { x = x << 1 | (i >> (bit * 2) & 1); y = y << 1 | (i >> (bit * 2 + 1) & 1); }
                x += bx; y += by;
                if (x < Width && y < Height) samples[y * Width + x] = plane.ReconstructionCoefficient(block * 1024 + i);
            }
        }
        for (int scale = 16; scale >= minimumScale; scale >>= 1) {
            for (int x = 0; x < Width; x += scale) {
                _budget.Cancellation.ThrowIfCancellationRequested();
                TransformLine(samples, x, Width * scale, (Height + scale - 1) / scale, false);
            }
            for (int y = 0; y < Height; y += scale) {
                _budget.Cancellation.ThrowIfCancellationRequested();
                TransformLine(samples, y * Width, scale, (Width + scale - 1) / scale, true);
            }
        }
        return samples;
    }

    private static void TransformLine(int[] values, int start, int stride, int count, bool horizontal) {
        int At(int i) => (uint)i < (uint)count ? values[start + i * stride] : 0;
        // Native IW44 short horizontal lines extend the final odd sample during
        // lifting. Longer lines and vertical lines use zero outside the image.
        // This compatibility edge is derived from independent 1-D/2-D renders;
        // it is separate from the prediction pass's nearest/average boundaries.
        int Lift(int i) => i >= count && horizontal && count >= 4 && count < 8
            ? At((count & 1) == 0 ? count - 1 : count - 2)
            : At(i);
        for (int i = 0; i < count; i += 2) values[start + i * stride] -= (9 * (Lift(i - 1) + Lift(i + 1)) - Lift(i - 3) - Lift(i + 3) + 16) >> 5;
        for (int i = 1; i < count; i += 2) {
            int prediction = i >= 3 && i + 3 < count ? (9 * (At(i - 1) + At(i + 1)) - At(i - 3) - At(i + 3) + 8) >> 4
                : i + 1 < count ? (At(i - 1) + At(i + 1) + 1) >> 1 : At(i - 1);
            values[start + i * stride] += prediction;
        }
    }
}
