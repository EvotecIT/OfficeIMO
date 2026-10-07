using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    // T.832 9.9.3 and 9.9.6 partition a raster into disjoint corner, edge,
    // and interior post-filter regions. Soft tiles share the raster boundary;
    // hard tiles apply the same partition to each tile independently.
    private static void ApplyOverlap(FrameHeader frame, int[] samples, int width, int height,
            int tileScale, long[] work, CancellationToken cancellation, int tileScaleY = 0) {
        if (!frame.HardTiles) {
            FilterRegion(samples, width, 0, 0, width, height, work, cancellation);
            return;
        }
        int top = 0;
        foreach (int tileHeight in frame.TileHeights) {
            int left = 0, bottom = top + tileHeight * (tileScaleY == 0 ? tileScale : tileScaleY);
            foreach (int tileWidth in frame.TileWidths) {
                int right = left + tileWidth * tileScale;
                FilterRegion(samples, width, left, top, right, bottom, work, cancellation);
                left = right;
            }
            top = bottom;
        }
    }

    private static void FilterRegion(int[] samples, int stride, int left, int top, int right, int bottom,
            long[] work, CancellationToken cancellation) {
        FilterCorner(samples, stride, left, top, work);
        FilterCorner(samples, stride, right - 2, top, work);
        FilterCorner(samples, stride, left, bottom - 2, work);
        FilterCorner(samples, stride, right - 2, bottom - 2, work);
        for (int x = left + 4; x < right; x += 4) {
            cancellation.ThrowIfCancellationRequested();
            for (int i = 0; i < 2; i++) {
                int upper = (top + i) * stride + x - 2, lower = (bottom - 2 + i) * stride + x - 2;
                FilterFour(samples, upper, upper + 1, upper + 2, upper + 3, work);
                FilterFour(samples, lower, lower + 1, lower + 2, lower + 3, work);
            }
        }
        for (int y = top + 4; y < bottom; y += 4) {
            cancellation.ThrowIfCancellationRequested();
            for (int i = 0; i < 2; i++) {
                int first = (y - 2) * stride + left + i, last = (y - 2) * stride + right - 2 + i;
                FilterFour(samples, first, first + stride, first + stride * 2, first + stride * 3, work);
                FilterFour(samples, last, last + stride, last + stride * 2, last + stride * 3, work);
            }
            for (int x = left + 4; x < right; x += 4) {
                int start = (y - 2) * stride + x - 2;
                for (int py = 0; py < 4; py++) for (int px = 0; px < 4; px++) work[py * 4 + px] = samples[start + py * stride + px];
                PostFilterSquare(work);
                for (int py = 0; py < 4; py++) for (int px = 0; px < 4; px++) samples[start + py * stride + px] = CheckedCoefficient(work[py * 4 + px]);
            }
        }
    }

    private static void FilterCorner(int[] samples, int stride, int x, int y, long[] work) {
        int start = y * stride + x;
        FilterFour(samples, start, start + 1, start + stride, start + stride + 1, work);
    }

    private static void FilterFour(int[] samples, int a, int b, int c, int d, long[] work) {
        work[0] = samples[a]; work[1] = samples[b]; work[2] = samples[c]; work[3] = samples[d];
        work[0] += work[3]; work[1] += work[2];
        work[3] -= (work[0] + 1) >> 1; work[2] -= (work[1] + 1) >> 1;
        InverseScale(work, 0, 3); InverseScale(work, 1, 2);
        work[0] += (work[3] * 3 + 4) >> 3; work[1] += (work[2] * 3 + 4) >> 3;
        work[3] -= work[0] >> 1; work[2] -= work[1] >> 1;
        work[0] += work[3]; work[1] += work[2];
        work[3] = -work[3]; work[2] = -work[2];
        InverseRotate(work, 2, 3);
        work[3] += (work[0] + 1) >> 1; work[2] += (work[1] + 1) >> 1;
        work[0] -= work[3]; work[1] -= work[2];
        samples[a] = CheckedCoefficient(work[0]); samples[b] = CheckedCoefficient(work[1]);
        samples[c] = CheckedCoefficient(work[2]); samples[d] = CheckedCoefficient(work[3]);
    }

    private static void PostFilterSquare(long[] work) {
        Hadamard(work, 0, 3, 12, 15, 0); Hadamard(work, 1, 2, 13, 14, 0);
        Hadamard(work, 4, 7, 8, 11, 0); Hadamard(work, 5, 6, 9, 10, 0);
        InverseRotate(work, 13, 12); InverseRotate(work, 9, 8);
        InverseRotate(work, 7, 3); InverseRotate(work, 6, 2);
        PostOddOdd(work, 10, 11, 14, 15);
        InverseScale(work, 0, 15); InverseScale(work, 1, 14);
        InverseScale(work, 4, 11); InverseScale(work, 5, 10);
        PostHadamard(work, 0, 3, 12, 15); PostHadamard(work, 1, 2, 13, 14);
        PostHadamard(work, 4, 7, 8, 11); PostHadamard(work, 5, 6, 9, 10);
    }

    private static void InverseRotate(long[] a, int first, int second) {
        a[first] -= (a[second] + 1) >> 1;
        a[second] += (a[first] + 1) >> 1;
    }

    private static void InverseScale(long[] a, int first, int second) {
        a[first] += a[second]; a[second] = (a[first] >> 1) - a[second];
        a[first] += (a[second] * 3) >> 3;
        a[second] += (a[first] * 3) >> 4;
        a[second] += a[first] >> 7; a[second] -= a[first] >> 10;
    }

    private static void PostOddOdd(long[] a, int i0, int i1, int i2, int i3) {
        a[i3] += a[i0]; a[i2] -= a[i1];
        long first = a[i3] >> 1, second = a[i2] >> 1;
        a[i0] -= first; a[i1] += second;
        a[i0] -= (a[i1] * 3 + 6) >> 3; a[i1] += (a[i0] * 3 + 2) >> 2; a[i0] -= (a[i1] * 3 + 4) >> 3;
        a[i1] -= second; a[i0] += first; a[i2] += a[i1]; a[i3] -= a[i0];
    }

    private static void PostHadamard(long[] a, int i0, int i1, int i2, int i3) {
        a[i1] -= a[i2]; a[i0] += (a[i3] * 3 + 4) >> 3;
        a[i3] -= a[i1] >> 1; a[i2] = ((a[i0] - a[i1]) >> 1) - a[i2];
        long swap = a[i2]; a[i2] = a[i3]; a[i3] = swap;
        a[i0] -= a[i3]; a[i1] += a[i2];
    }
}
