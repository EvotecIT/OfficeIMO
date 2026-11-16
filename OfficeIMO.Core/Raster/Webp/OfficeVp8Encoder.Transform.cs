// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// OfficeIMO's copy is licensed under the repository's MIT license by the original author.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    // RFC 6386 section 14.3 reconstructs a 4x4 block as X = A^T * Y * A / 8, where the rows of A are
    // (1, 1, 1, 1), (c, s, -s, -c), (1, -1, -1, 1) and (s, -c, c, -s), with c = sqrt(2) * cos(pi / 8) and
    // s = sqrt(2) * sin(pi / 8). The rows are orthogonal with squared length 4, so the exact forward transform
    // is Y = A * X * A^T / 2. Coefficient Y[v, u] is stored at 4 * v + u, matching the decoder's layout.
    private static readonly double CosineBasis = Math.Sqrt(2D) * Math.Cos(Math.PI / 8D);
    private static readonly double SineBasis = Math.Sqrt(2D) * Math.Sin(Math.PI / 8D);

    private void ComputeCoefficients(int[] residual, double[] coefficients) => ForwardTransform(residual, _transformTemporaryDouble, coefficients);

    internal static void ForwardTransform(int[] residual, double[] rows, double[] coefficients) {
        for (int row = 0; row < 4; row++) {
            int at = row * 4;
            ForwardBasis(residual[at], residual[at + 1], residual[at + 2], residual[at + 3], rows, at, 1);
        }

        for (int column = 0; column < 4; column++) {
            ForwardBasis(rows[column], rows[column + 4], rows[column + 8], rows[column + 12], coefficients, column, 4);
        }

        // The decoder receives integers; rounding each coefficient keeps the encoder's integer contract.
        for (int index = 0; index < 16; index++) {
            coefficients[index] = Math.Round(coefficients[index] / 2D, MidpointRounding.AwayFromZero);
        }
    }

    private static void ForwardBasis(double x0, double x1, double x2, double x3, double[] output, int offset, int stride) {
        double sumOuter = x0 + x3, sumInner = x1 + x2;
        double differenceOuter = x0 - x3, differenceInner = x1 - x2;
        output[offset] = sumOuter + sumInner;
        output[offset + stride] = CosineBasis * differenceOuter + SineBasis * differenceInner;
        output[offset + 2 * stride] = sumOuter - sumInner;
        output[offset + 3 * stride] = SineBasis * differenceOuter - CosineBasis * differenceInner;
    }

    private void ComputeWalshCoefficients(double[] dcValues, double[] coefficients) {
        double[] temporary = _transformTemporaryDouble;
        for (int row = 0; row < 4; row++) {
            int at = row * 4;
            double a = dcValues[at] + dcValues[at + 3], b = dcValues[at + 1] + dcValues[at + 2];
            double c = dcValues[at + 1] - dcValues[at + 2], d = dcValues[at] - dcValues[at + 3];
            temporary[at] = a + b; temporary[at + 1] = d + c;
            temporary[at + 2] = a - b; temporary[at + 3] = d - c;
        }
        for (int col = 0; col < 4; col++) {
            double a = temporary[col] + temporary[col + 12], b = temporary[col + 4] + temporary[col + 8];
            double c = temporary[col + 4] - temporary[col + 8], d = temporary[col] - temporary[col + 12];
            coefficients[col] = (a + b) / 2; coefficients[col + 4] = (d + c) / 2;
            coefficients[col + 8] = (a - b) / 2; coefficients[col + 12] = (d - c) / 2;
        }
    }
}
