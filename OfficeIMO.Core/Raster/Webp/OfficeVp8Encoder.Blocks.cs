// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    private void EncodeLumaMacroblockY16(
        OfficeVp8BoolEncoder tokenWriter,
        int[] probabilities,
        byte[] yPlane,
        byte[] reconY,
        int width,
        int height,
        int mbX,
        int mbY,
        int mode,
        DequantFactors dequant,
        byte[] y2NzAbove,
        byte[] y2NzCurrent,
        ref byte y2Left,
        byte[] yNzAbove,
        byte[] yNzCurrent) {
        var macroblockOffsetX = mbX * MacroblockSize;
        var macroblockOffsetY = mbY * MacroblockSize;
        var yBlockBase = mbX * MacroblockSubBlockCount;

        byte[] macroPrediction = _blocksMacroPrediction;
        OfficeVp8Prediction.PredictBlock(reconY, width, height, macroblockOffsetX, macroblockOffsetY, 16, mode, macroPrediction, _scratch);
        var dcValues = _blocksDcValues;
        var yQuant = _blocksYQuant;

        byte[] predicted = _blocksPredicted;
        int[] residual = _blocksResidual;
        double[] coeffs = _blocksCoeffs;

        for (var blockIndex = 0; blockIndex < MacroblockSubBlockCount; blockIndex++) {
            var subX = blockIndex & 3;
            var subY = blockIndex >> 2;
            var dstX = macroblockOffsetX + (subX * BlockSize);
            var dstY = macroblockOffsetY + (subY * BlockSize);

            CopyPredictionSubblock(macroPrediction, 16, subX * 4, subY * 4, predicted);

            FillResidual(yPlane, width, height, dstX, dstY, predicted, residual);

            ComputeCoefficients(residual, coeffs);

            dcValues[blockIndex] = coeffs[0];
            var offset = blockIndex * CoefficientsPerBlock;
            yQuant[offset] = 0;
            for (var i = 1; i < CoefficientsPerBlock; i++) {
                yQuant[offset + i] = ClampCoefficient(QuantizeDouble(coeffs[i], dequant.Y1Ac));
            }

            var qdc = ClampCoefficient(QuantizeDouble(coeffs[0], dequant.Y1Dc));
            var dequantCoeffs = _blocksDequantCoeffs;
            dequantCoeffs[0] = qdc * dequant.Y1Dc;
            for (var i = 1; i < CoefficientsPerBlock; i++) {
                dequantCoeffs[i] = yQuant[offset + i] * dequant.Y1Ac;
            }

            var residualDecoded = ReconstructBlock(dequantCoeffs);
            UpdateReconstruction(reconY, width, height, dstX, dstY, predicted, residualDecoded);
        }

        double[] y2Coeff = _blocksY2Coeff;
        ComputeWalshCoefficients(dcValues, y2Coeff);

        var y2Quant = _blocksY2Quant;
        for (var i = 0; i < CoefficientsPerBlock; i++) {
            var dequantFactor = i == 0 ? dequant.Y2Dc : dequant.Y2Ac;
            y2Quant[i] = ClampCoefficient(QuantizeDouble(y2Coeff[i], dequantFactor));
        }

        var y2InitialContext = 0;
        if (mbY > 0) y2InitialContext += y2NzAbove[mbX];
        y2InitialContext += y2Left;
        if (y2InitialContext > 2) y2InitialContext = 2;

        var hasNonZeroY2 = EncodeBlockCoefficients(
            tokenWriter,
            probabilities,
            BlockTypeY2,
            y2InitialContext,
            y2Quant);
        y2Left = hasNonZeroY2 ? (byte)1 : (byte)0;
        y2NzCurrent[mbX] = y2Left;

        for (var blockIndex = 0; blockIndex < MacroblockSubBlockCount; blockIndex++) {
            var subX = blockIndex & 3;
            var subY = blockIndex >> 2;
            var initialContext = 0;

            if (subY == 0) {
                if (mbY > 0) initialContext += yNzAbove[yBlockBase + 12 + subX];
            } else {
                initialContext += yNzCurrent[yBlockBase + ((subY - 1) * 4) + subX];
            }

            if (subX == 0) {
                if (mbX > 0) initialContext += yNzCurrent[yBlockBase - MacroblockSubBlockCount + (subY * 4) + 3];
            } else {
                initialContext += yNzCurrent[yBlockBase + (subY * 4) + subX - 1];
            }

            if (initialContext > 2) initialContext = 2;

            var offset = blockIndex * CoefficientsPerBlock;
            var coeffTokens = _blocksCoeffTokens;
            coeffTokens[0] = 0;
            for (var i = 1; i < CoefficientsPerBlock; i++) {
                coeffTokens[i] = yQuant[offset + i];
            }

            var hasNonZero = EncodeBlockCoefficients(
                tokenWriter,
                probabilities,
                OfficeVp8Tables.LumaAc,
                initialContext,
                coeffTokens);

            yNzCurrent[yBlockBase + blockIndex] = hasNonZero ? (byte)1 : (byte)0;
        }

        var y2Dequant = _blocksY2Dequant;
        for (var i = 0; i < CoefficientsPerBlock; i++) {
            var dequantFactor = i == 0 ? dequant.Y2Dc : dequant.Y2Ac;
            y2Dequant[i] = y2Quant[i] * dequantFactor;
        }

        var dcOverride = OfficeVp8Transform.InverseWalshTransform4x4(y2Dequant);
        for (var blockIndex = 0; blockIndex < MacroblockSubBlockCount; blockIndex++) {
            var subX = blockIndex & 3;
            var subY = blockIndex >> 2;
            var dstX = macroblockOffsetX + (subX * BlockSize);
            var dstY = macroblockOffsetY + (subY * BlockSize);

            CopyPredictionSubblock(macroPrediction, 16, subX * 4, subY * 4, predicted);

            var offset = blockIndex * CoefficientsPerBlock;
            var dequantCoeffs = _blocksDequantCoeffs;
            dequantCoeffs[0] = dcOverride[blockIndex];
            for (var i = 1; i < CoefficientsPerBlock; i++) {
                dequantCoeffs[i] = yQuant[offset + i] * dequant.Y1Ac;
            }

            var residualDecoded = ReconstructBlock(dequantCoeffs);
            UpdateReconstruction(reconY, width, height, dstX, dstY, predicted, residualDecoded);
        }
    }

    private bool EncodeBlock(
        OfficeVp8BoolEncoder tokenWriter,
        int[] probabilities,
        int blockType,
        int initialContext,
        byte[] source,
        byte[] recon,
        int planeWidth,
        int planeHeight,
        int dstX,
        int dstY,
        int mode,
        int dequantDc,
        int dequantAc) {
        byte[] predicted = _blocksPredicted;
        if (blockType == BlockTypeU) {
            for (int y = 0; y < 4; y++) for (int x = 0; x < 4; x++) predicted[y * 4 + x] = recon[(dstY + y) * planeWidth + dstX + x];
        } else {
            PredictBlock(recon, planeWidth, planeHeight, dstX, dstY, mode, predicted);
        }

        int[] residual = _blocksResidual;
        FillResidual(source, planeWidth, planeHeight, dstX, dstY, predicted, residual);

        double[] dequantized = _blocksDequantized;
        ComputeCoefficients(residual, dequantized);

        var coefficients = _blocksCoefficients;
        for (var i = 0; i < CoefficientsPerBlock; i++) {
            var dequant = i == 0 ? dequantDc : dequantAc;
            coefficients[i] = ClampCoefficient(QuantizeDouble(dequantized[i], dequant));
        }

        var hasNonZero = EncodeBlockCoefficients(
            tokenWriter,
            probabilities,
            blockType,
            initialContext,
            coefficients);

        var dequantCoeffs = _blocksDequantCoeffs;
        for (var i = 0; i < CoefficientsPerBlock; i++) {
            var dequant = i == 0 ? dequantDc : dequantAc;
            dequantCoeffs[i] = coefficients[i] * dequant;
        }

        var residualDecoded = ReconstructBlock(dequantCoeffs);
        UpdateReconstruction(recon, planeWidth, planeHeight, dstX, dstY, predicted, residualDecoded);

        return hasNonZero;
    }

    private int[] ReconstructBlock(int[] coefficients) {
        OfficeVp8Transform.InverseTransform4x4(coefficients, _scratch.TransformTemp, _scratch.TransformOutput);
        return _scratch.TransformOutput;
    }

}
