// Adapted from CodeGlyphX, commit fc25e2fcf795d9c9a09b88708c47bdfeed5c446d.
// OfficeIMO's copy is licensed under the repository's MIT license by the original author.
// OfficeIMO adaptation removes diagnostic scaffolding and adds bounded cancellation/resource handling.
using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeVp8Decoder {
    private static void ApplyDecodedResidual(byte[] plane, int width, int height, int x, int y,
        int[] coefficients, bool is4x4, int mode, bool overrideDc, int dcValue, OfficeVp8DecodeScratch scratch) {
        if (is4x4) PredictDecodedSubblock(plane, width, height, x, y, mode, scratch);
        int originalDc = coefficients[0];
        if (overrideDc) coefficients[0] = dcValue;
        OfficeVp8Transform.InverseTransform4x4(coefficients, scratch.TransformTemp, scratch.TransformOutput);
        if (overrideDc) coefficients[0] = originalDc;
        int[] residual = scratch.TransformOutput;
        for (int row = 0; row < 4 && y + row < height; row++) {
            for (int col = 0; col < 4 && x + col < width; col++) {
                int index = (y + row) * width + x + col;
                plane[index] = ClampToByte(plane[index] + residual[row * 4 + col]);
            }
        }
    }

    private static void PredictDecodedBlock(byte[] plane, int width, int height, int x, int y, int size, int mode, OfficeVp8DecodeScratch scratch) {
        byte[] predicted = scratch.Predicted;
        OfficeVp8Prediction.PredictBlock(plane, width, height, x, y, size, mode, predicted, scratch);
        CopyDecodedPrediction(plane, width, height, x, y, size, predicted);
    }
    private static void PredictDecodedSubblock(byte[] plane, int width, int height, int x, int y, int mode, OfficeVp8DecodeScratch scratch) {
        byte[] predicted = scratch.Predicted;
        OfficeVp8Prediction.PredictSubblock(plane, width, height, x, y, mode, predicted, scratch);
        CopyDecodedPrediction(plane, width, height, x, y, 4, predicted);
    }
    private static void CopyDecodedPrediction(byte[] plane, int width, int height, int x, int y, int size, byte[] predicted) {
        for (int row = 0; row < size && y + row < height; row++) {
            for (int col = 0; col < size && x + col < width; col++) {
                plane[(y + row) * width + x + col] = predicted[row * size + col];
            }
        }
    }
}
