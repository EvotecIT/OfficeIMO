using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    private static void ApplyReducedOverlap(FrameHeader frame, int[] samples, int width, int height,
            int tileScaleY, CancellationToken cancellation) {
        if (!frame.HardTiles) {
            FilterReducedRegion(samples, width, 0, 0, width, height, cancellation);
            return;
        }
        int top = 0;
        foreach (int tileHeight in frame.TileHeights) {
            int left = 0, bottom = top + tileHeight * tileScaleY;
            foreach (int tileWidth in frame.TileWidths) {
                int right = left + tileWidth * 2;
                FilterReducedRegion(samples, width, left, top, right, bottom, cancellation);
                left = right;
            }
            top = bottom;
        }
    }

    // T.832 9.9.3.3-4: both reduced formats share a 2x2 overlap lattice.
    private static void FilterReducedRegion(int[] a, int stride, int left, int top, int right, int bottom,
            CancellationToken cancellation) {
        AdjustReducedCorners(a, stride, left, top, right, bottom, -1);
        for (int x = left + 2; x < right; x += 2) {
            cancellation.ThrowIfCancellationRequested();
            FilterTwo(a, top * stride + x - 1, top * stride + x);
            FilterTwo(a, (bottom - 1) * stride + x - 1, (bottom - 1) * stride + x);
        }
        for (int y = top + 2; y < bottom; y += 2) {
            cancellation.ThrowIfCancellationRequested();
            FilterTwo(a, (y - 1) * stride + left, y * stride + left);
            FilterTwo(a, (y - 1) * stride + right - 1, y * stride + right - 1);
            for (int x = left + 2; x < right; x += 2) {
                int first = (y - 1) * stride + x - 1;
                FilterTwoSquare(a, first, first + 1, first + stride, first + stride + 1);
            }
        }
        AdjustReducedCorners(a, stride, left, top, right, bottom, 1);
    }

    private static void AdjustReducedCorners(int[] a, int stride, int left, int top, int right, int bottom, int sign) {
        int upper = top * stride, lower = (bottom - 1) * stride;
        a[upper + left] = CheckedCoefficient((long)a[upper + left] + sign * (long)a[upper + left + 1]);
        a[upper + right - 1] = CheckedCoefficient((long)a[upper + right - 1] + sign * (long)a[upper + right - 2]);
        a[lower + left] = CheckedCoefficient((long)a[lower + left] + sign * (long)a[lower + left + 1]);
        a[lower + right - 1] = CheckedCoefficient((long)a[lower + right - 1] + sign * (long)a[lower + right - 2]);
    }

    private static void FilterTwo(int[] samples, int first, int second) {
        long a = samples[first], b = samples[second];
        PostFilterTwo(ref a, ref b);
        samples[first] = CheckedCoefficient(a); samples[second] = CheckedCoefficient(b);
    }

    private static void PostFilterTwo(ref long a, ref long b) {
        b += (a + 2) >> 2;
        a += (b + 1) >> 1;
        a += b >> 5; a += b >> 9; a += b >> 13;
        b += (a + 2) >> 2;
    }

    private static void FilterTwoSquare(int[] samples, int i0, int i1, int i2, int i3) {
        long a = samples[i0], b = samples[i1], c = samples[i2], d = samples[i3];
        a += d; b += c;
        d -= (a + 1) >> 1; c -= (b + 1) >> 1;
        PostFilterTwo(ref a, ref b);
        d += (a + 1) >> 1; c += (b + 1) >> 1;
        a -= d; b -= c;
        samples[i0] = CheckedCoefficient(a); samples[i1] = CheckedCoefficient(b);
        samples[i2] = CheckedCoefficient(c); samples[i3] = CheckedCoefficient(d);
    }
}
