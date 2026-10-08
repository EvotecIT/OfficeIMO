// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    // Forward transform adapted from WebM libvpx (BSD 3-Clause); see Licenses/libvpx-LICENSE.txt and libvpx-PATENTS.txt.
    // VP8 forward integer transform. Keep the specified rounding at each pass;
    // inverting a rounded unit-basis IDCT produces a singular matrix.
    private void ComputeCoefficients(int[] residual, double[] coefficients) {
        int[] temporary = _transformTemporaryInt;
        for (int row = 0; row < 4; row++) {
            int at = row * 4;
            int a = (residual[at] + residual[at + 3]) * 8;
            int b = (residual[at + 1] + residual[at + 2]) * 8;
            int c = (residual[at + 1] - residual[at + 2]) * 8;
            int d = (residual[at] - residual[at + 3]) * 8;
            temporary[at] = a + b;
            temporary[at + 2] = a - b;
            temporary[at + 1] = (c * 2217 + d * 5352 + 14500) >> 12;
            temporary[at + 3] = (d * 2217 - c * 5352 + 7500) >> 12;
        }
        for (int col = 0; col < 4; col++) {
            int a = temporary[col] + temporary[col + 12];
            int b = temporary[col + 4] + temporary[col + 8];
            int c = temporary[col + 4] - temporary[col + 8];
            int d = temporary[col] - temporary[col + 12];
            coefficients[col] = (a + b + 7) >> 4;
            coefficients[col + 8] = (a - b + 7) >> 4;
            coefficients[col + 4] = ((c * 2217 + d * 5352 + 12000) >> 16) + (d == 0 ? 0 : 1);
            coefficients[col + 12] = (d * 2217 - c * 5352 + 51000) >> 16;
        }
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
