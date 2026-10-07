using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    private static readonly int[] InversePermutation = { 0, 8, 4, 13, 2, 15, 3, 14, 1, 12, 5, 9, 7, 11, 6, 10 };

    // T.832 9.9.7: the reversible integer lifting transform. Wider intermediates
    // prevent hostile coefficients from wrapping before the range check.
    private static void InverseTransform(int[] coefficients, int offset, long[] work) {
        for (int i = 0; i < 16; i++) work[InversePermutation[i]] = coefficients[offset + i];
        Hadamard(work, 0, 1, 4, 5, 1);
        Odd(work, 2, 3, 6, 7);
        Odd(work, 8, 12, 9, 13);
        OddOdd(work, 10, 11, 14, 15);
        Hadamard(work, 0, 3, 12, 15, 0);
        Hadamard(work, 5, 6, 9, 10, 0);
        Hadamard(work, 1, 2, 13, 14, 0);
        Hadamard(work, 4, 7, 8, 11, 0);
        for (int i = 0; i < 16; i++) coefficients[offset + i] = CheckedCoefficient(work[i]);
    }

    private static void Hadamard(long[] a, int i0, int i1, int i2, int i3, int round) {
        a[i0] += a[i3]; a[i1] -= a[i2];
        long first = (a[i0] - a[i1] + round) >> 1, second = a[i2];
        a[i2] = first - a[i3]; a[i3] = first - second;
        a[i0] -= a[i3]; a[i1] += a[i2];
    }

    private static void Odd(long[] a, int i0, int i1, int i2, int i3) {
        a[i1] += a[i3]; a[i0] -= a[i2];
        a[i3] -= a[i1] >> 1; a[i2] += (a[i0] + 1) >> 1;
        a[i0] -= (3 * a[i1] + 4) >> 3; a[i1] += (3 * a[i0] + 4) >> 3;
        a[i2] -= (3 * a[i3] + 4) >> 3; a[i3] += (3 * a[i2] + 4) >> 3;
        a[i2] -= (a[i1] + 1) >> 1; a[i3] = ((a[i0] + 1) >> 1) - a[i3];
        a[i1] += a[i2]; a[i0] -= a[i3];
    }

    private static void OddOdd(long[] a, int i0, int i1, int i2, int i3) {
        a[i3] += a[i0]; a[i2] -= a[i1];
        long first = a[i3] >> 1, second = a[i2] >> 1;
        a[i0] -= first; a[i1] += second;
        a[i0] -= (3 * a[i1] + 3) >> 3; a[i1] += (3 * a[i0] + 3) >> 2; a[i0] -= (3 * a[i1] + 4) >> 3;
        a[i1] -= second; a[i0] += first;
        a[i2] += a[i1]; a[i3] -= a[i0];
        a[i1] = -a[i1]; a[i2] = -a[i2];
    }
}
