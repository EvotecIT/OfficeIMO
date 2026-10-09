// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    private byte[] PadPlane(byte[] input, int width, int height, int paddedWidth, int paddedHeight) {
        var output = new byte[checked(paddedWidth * paddedHeight)];
        for (int y = 0; y < paddedHeight; y++) {
            Checkpoint();
            for (int x = 0; x < paddedWidth; x++) {
                output[y * paddedWidth + x] = input[Math.Min(y, height - 1) * width + Math.Min(x, width - 1)];
            }
        }
        return output;
    }
    private void PrefillPrediction(byte[] plane, int width, int height, int x, int y, int size, int mode) {
        byte[] predicted = _predictionPredicted;
        OfficeVp8Prediction.PredictBlock(plane, width, height, x, y, size, mode, predicted, _scratch);
        for (int row = 0; row < size; row++) for (int col = 0; col < size; col++)
            plane[(y + row) * width + x + col] = predicted[row * size + col];
    }
    private void CopyPredictionSubblock(byte[] prediction, int stride, int x, int y, byte[] block) {
        for (int row = 0; row < 4; row++) for (int col = 0; col < 4; col++)
            block[row * 4 + col] = prediction[(y + row) * stride + x + col];
    }
    private bool PlaneBlockMatches(byte[] source, byte[] prediction, int width, int x, int y, int size) {
        for (int row = 0; row < size; row++) for (int col = 0; col < size; col++) {
            int index = (y + row) * width + x + col;
            if (source[index] != prediction[index]) return false;
        }
        return true;
    }
}
