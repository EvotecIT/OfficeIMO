// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    private int ChooseBestBlockMode(
        byte[] source,
        byte[] recon,
        int planeWidth,
        int planeHeight,
        int dstX,
        int dstY,
        out int bestCost) {
        var bestMode = ModeDcPred;
        bestCost = int.MaxValue;
        byte[] predicted = _predictionModesPredicted;

        for (var mode = 0; mode < Intra4x4ModeCount; mode++) {
            PredictBlock(recon, planeWidth, planeHeight, dstX, dstY, mode, predicted);
            var cost = ComputePredictionCost(source, planeWidth, planeHeight, dstX, dstY, predicted);
            if (cost < bestCost) {
                bestCost = cost;
                bestMode = mode;
            }
        }

        return bestMode;
    }

    private int ChooseBestUvMode(
        byte[] uPlane,
        byte[] vPlane,
        byte[] reconU,
        byte[] reconV,
        int chromaWidth,
        int chromaHeight,
        int chromaOffsetX,
        int chromaOffsetY) {
        var bestMode = ModeDcPred;
        var bestCost = long.MaxValue;
        byte[] predictedU = _predictionModesPredictedU;
        byte[] predictedV = _predictionModesPredictedV;
        for (var mode = 0; mode <= 3; mode++) {
            OfficeVp8Prediction.PredictBlock(reconU, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, 8, mode, predictedU, _scratch);
            OfficeVp8Prediction.PredictBlock(reconV, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, 8, mode, predictedV, _scratch);
            long cost = 0;
            for (int y = 0; y < 8; y++) for (int x = 0; x < 8; x++) {
                int index = (chromaOffsetY + y) * chromaWidth + chromaOffsetX + x;
                cost += Math.Abs(uPlane[index] - predictedU[y * 8 + x]);
                cost += Math.Abs(vPlane[index] - predictedV[y * 8 + x]);
            }
            if (cost < bestCost) { bestCost = cost; bestMode = mode; }
        }

        return bestMode;
    }

    private int ComputePredictionCost(
        byte[] source,
        int planeWidth,
        int planeHeight,
        int dstX,
        int dstY,
        byte[] predicted) {
        var sum = 0;
        for (var y = 0; y < BlockSize; y++) {
            Checkpoint();
            var py = dstY + y;
            if ((uint)py >= (uint)planeHeight) continue;
            var rowOffset = py * planeWidth;
            var predRow = y * BlockSize;
            for (var x = 0; x < BlockSize; x++) {
                var px = dstX + x;
                if ((uint)px >= (uint)planeWidth) continue;
                var diff = source[rowOffset + px] - predicted[predRow + x];
                if (diff < 0) diff = -diff;
                sum += diff;
            }
        }

        return sum;
    }

    private bool ApplyPredictionAndCheckMatchY16(
        byte[] yPlane,
        byte[] uPlane,
        byte[] vPlane,
        byte[] reconY,
        byte[] reconU,
        byte[] reconV,
        int width,
        int height,
        int chromaWidth,
        int chromaHeight,
        int macroblockOffsetX,
        int macroblockOffsetY,
        int chromaOffsetX,
        int chromaOffsetY,
        int yMode,
        int uvMode) {
        PrefillPrediction(reconY, width, height, macroblockOffsetX, macroblockOffsetY, 16, yMode);
        if (!PlaneBlockMatches(yPlane, reconY, width, macroblockOffsetX, macroblockOffsetY, 16)) return false;

        return ApplyPredictionAndCheckMatchUv(
            uPlane,
            vPlane,
            reconU,
            reconV,
            chromaWidth,
            chromaHeight,
            chromaOffsetX,
            chromaOffsetY,
            uvMode);
    }

    private bool ApplyPredictionAndCheckMatchBPred(
        byte[] yPlane,
        byte[] uPlane,
        byte[] vPlane,
        byte[] reconY,
        byte[] reconU,
        byte[] reconV,
        int width,
        int height,
        int chromaWidth,
        int chromaHeight,
        int macroblockOffsetX,
        int macroblockOffsetY,
        int chromaOffsetX,
        int chromaOffsetY,
        int uvMode,
        int[] outputModes) {
        byte[] predicted = _predictionModesPredicted;

        for (var blockIndex = 0; blockIndex < MacroblockSubBlockCount; blockIndex++) {
            var subX = blockIndex & 3;
            var subY = blockIndex >> 2;
            var dstX = macroblockOffsetX + (subX * BlockSize);
            var dstY = macroblockOffsetY + (subY * BlockSize);

            var mode = ChooseBestBlockMode(yPlane, reconY, width, height, dstX, dstY, out _);
            outputModes[blockIndex] = mode;

            PredictBlock(reconY, width, height, dstX, dstY, mode, predicted);
            if (!BlockMatchesPrediction(yPlane, width, height, dstX, dstY, predicted)) {
                return false;
            }

            CopyPredictedBlock(reconY, width, height, dstX, dstY, predicted);
        }

        return ApplyPredictionAndCheckMatchUv(
            uPlane,
            vPlane,
            reconU,
            reconV,
            chromaWidth,
            chromaHeight,
            chromaOffsetX,
            chromaOffsetY,
            uvMode);
    }

    private bool ApplyPredictionAndCheckMatchUv(
        byte[] uPlane,
        byte[] vPlane,
        byte[] reconU,
        byte[] reconV,
        int chromaWidth,
        int chromaHeight,
        int chromaOffsetX,
        int chromaOffsetY,
        int uvMode) {
        PrefillPrediction(reconU, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, 8, uvMode);
        PrefillPrediction(reconV, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, 8, uvMode);
        return PlaneBlockMatches(uPlane, reconU, chromaWidth, chromaOffsetX, chromaOffsetY, 8)
            && PlaneBlockMatches(vPlane, reconV, chromaWidth, chromaOffsetX, chromaOffsetY, 8);
    }

    private bool BlockMatchesPrediction(
        byte[] source,
        int planeWidth,
        int planeHeight,
        int dstX,
        int dstY,
        byte[] predicted) {
        for (var y = 0; y < BlockSize; y++) {
            Checkpoint();
            var py = dstY + y;
            if ((uint)py >= (uint)planeHeight) continue;
            var rowOffset = py * planeWidth;
            var predRow = y * BlockSize;
            for (var x = 0; x < BlockSize; x++) {
                var px = dstX + x;
                if ((uint)px >= (uint)planeWidth) continue;
                if (source[rowOffset + px] != predicted[predRow + x]) return false;
            }
        }

        return true;
    }

    private void CopyPredictedBlock(
        byte[] plane,
        int planeWidth,
        int planeHeight,
        int dstX,
        int dstY,
        byte[] predicted) {
        for (var y = 0; y < BlockSize; y++) {
            Checkpoint();
            var py = dstY + y;
            if ((uint)py >= (uint)planeHeight) continue;
            var rowOffset = py * planeWidth;
            var predRow = y * BlockSize;
            for (var x = 0; x < BlockSize; x++) {
                var px = dstX + x;
                if ((uint)px >= (uint)planeWidth) continue;
                plane[rowOffset + px] = predicted[predRow + x];
            }
        }
    }

    private void CopyPlaneBlock(
        byte[] plane,
        int planeWidth,
        int planeHeight,
        int dstX,
        int dstY,
        int blockWidth,
        int blockHeight,
        byte[] buffer) {
        Array.Clear(buffer, 0, buffer.Length);
        for (var y = 0; y < blockHeight; y++) {
            Checkpoint();
            var py = dstY + y;
            if ((uint)py >= (uint)planeHeight) break;
            var rowOffset = py * planeWidth;
            var bufferRow = y * blockWidth;
            for (var x = 0; x < blockWidth; x++) {
                var px = dstX + x;
                if ((uint)px >= (uint)planeWidth) break;
                buffer[bufferRow + x] = plane[rowOffset + px];
            }
        }
    }

    private void RestorePlaneBlock(
        byte[] plane,
        int planeWidth,
        int planeHeight,
        int dstX,
        int dstY,
        int blockWidth,
        int blockHeight,
        byte[] buffer) {
        for (var y = 0; y < blockHeight; y++) {
            Checkpoint();
            var py = dstY + y;
            if ((uint)py >= (uint)planeHeight) break;
            var rowOffset = py * planeWidth;
            var bufferRow = y * blockWidth;
            for (var x = 0; x < blockWidth; x++) {
                var px = dstX + x;
                if ((uint)px >= (uint)planeWidth) break;
                plane[rowOffset + px] = buffer[bufferRow + x];
            }
        }
    }

    private void PredictBlock(byte[] plane, int planeWidth, int planeHeight, int dstX, int dstY, int mode, byte[] predicted) {
        OfficeVp8Prediction.PredictSubblock(plane, planeWidth, planeHeight, dstX, dstY, mode, predicted, _scratch);
    }

    private void FillResidual(
        byte[] source,
        int planeWidth,
        int planeHeight,
        int dstX,
        int dstY,
        byte[] predicted,
        int[] residual) {
        Array.Clear(residual, 0, residual.Length);
        for (var y = 0; y < BlockSize; y++) {
            Checkpoint();
            var py = dstY + y;
            if ((uint)py >= (uint)planeHeight) continue;
            var rowOffset = py * planeWidth;
            var predRow = y * BlockSize;
            var resRow = y * BlockSize;
            for (var x = 0; x < BlockSize; x++) {
                var px = dstX + x;
                if ((uint)px >= (uint)planeWidth) continue;
                residual[resRow + x] = source[rowOffset + px] - predicted[predRow + x];
            }
        }
    }

    private void UpdateReconstruction(
        byte[] recon,
        int planeWidth,
        int planeHeight,
        int dstX,
        int dstY,
        byte[] predicted,
        int[] residual) {
        for (var y = 0; y < BlockSize; y++) {
            Checkpoint();
            var py = dstY + y;
            if ((uint)py >= (uint)planeHeight) continue;
            var rowOffset = py * planeWidth;
            var predRow = y * BlockSize;
            var resRow = y * BlockSize;
            for (var x = 0; x < BlockSize; x++) {
                var px = dstX + x;
                if ((uint)px >= (uint)planeWidth) continue;
                recon[rowOffset + px] = ClampToByte(predicted[predRow + x] + residual[resRow + x]);
            }
        }
    }

}
