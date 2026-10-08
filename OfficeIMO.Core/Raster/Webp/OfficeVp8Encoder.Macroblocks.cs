// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    private void EncodeMacroblock(
        OfficeVp8BoolEncoder headerWriter,
        OfficeVp8BoolEncoder tokenWriter,
        int[] probabilities,
        byte[] yPlane,
        byte[] uPlane,
        byte[] vPlane,
        byte[] reconY,
        byte[] reconU,
        byte[] reconV,
        int width,
        int height,
        int mbX,
        int mbY,
        DequantFactors[] dequantSegments,
        int segmentId,
        byte[] y2NzAbove,
        byte[] y2NzCurrent,
        ref byte y2Left,
        byte[] yNzAbove,
        byte[] yNzCurrent,
        byte[] uNzAbove,
        byte[] uNzCurrent,
        byte[] vNzAbove,
        byte[] vNzCurrent,
        int[] aboveSubModes,
        int[] currentSubModes,
        bool segmentationEnabled,
        int[] segmentProbabilities,
        bool enableSkip,
        int skipProbability) {
        var macroblockOffsetX = mbX * MacroblockSize;
        var macroblockOffsetY = mbY * MacroblockSize;
        var chromaWidth = (width + 1) >> 1;
        var chromaHeight = (height + 1) >> 1;
        var chromaOffsetX = mbX * (MacroblockSize / 2);
        var chromaOffsetY = mbY * (MacroblockSize / 2);

        if (segmentationEnabled) {
            WriteSegmentId(headerWriter, segmentProbabilities, segmentId);
        }

        if ((uint)segmentId >= SegmentCount) segmentId = 0;
        var dequant = (dequantSegments != null && dequantSegments.Length > segmentId)
            ? dequantSegments[segmentId]
            : dequantSegments?[0] ?? BuildDequantFactors(0);

        var yBlockBase = mbX * MacroblockSubBlockCount;

        long bestBPredCost = 0;
        for (var blockIndex = 0; blockIndex < MacroblockSubBlockCount; blockIndex++) {
            var subX = blockIndex & 3;
            var subY = blockIndex >> 2;
            var dstX = macroblockOffsetX + (subX * BlockSize);
            var dstY = macroblockOffsetY + (subY * BlockSize);
            ChooseBestBlockMode(yPlane, reconY, width, height, dstX, dstY, out var cost);
            bestBPredCost += cost;
        }

        var bestY16Mode = ModeDcPred;
        long bestY16Cost = long.MaxValue;
        byte[] predictedMacroblock = _macroblocksPredictedMacroblock;
        for (var mode = 0; mode <= 3; mode++) {
            OfficeVp8Prediction.PredictBlock(reconY, width, height, macroblockOffsetX, macroblockOffsetY, 16, mode, predictedMacroblock, _scratch);
            long cost = 0;
            for (int y = 0; y < 16; y++) for (int x = 0; x < 16; x++)
                cost += Math.Abs(yPlane[(macroblockOffsetY + y) * width + macroblockOffsetX + x] - predictedMacroblock[y * 16 + x]);
            if (cost < bestY16Cost) { bestY16Cost = cost; bestY16Mode = mode; }
        }

        var useY16 = bestY16Cost <= bestBPredCost;
        var uvMode = ChooseBestUvMode(
            uPlane,
            vPlane,
            reconU,
            reconV,
            chromaWidth,
            chromaHeight,
            chromaOffsetX,
            chromaOffsetY);

        var skipCoefficients = false;
        int[]? bPredModes = null;
        byte[] yBackup = _macroblocksYBackup;
        byte[] uBackup = _macroblocksUBackup;
        byte[] vBackup = _macroblocksVBackup;
        if (enableSkip) {
            CopyPlaneBlock(reconY, width, height, macroblockOffsetX, macroblockOffsetY, MacroblockSize, MacroblockSize, yBackup);
            CopyPlaneBlock(reconU, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, MacroblockSize / 2, MacroblockSize / 2, uBackup);
            CopyPlaneBlock(reconV, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, MacroblockSize / 2, MacroblockSize / 2, vBackup);

            if (useY16) {
                skipCoefficients = ApplyPredictionAndCheckMatchY16(
                    yPlane,
                    uPlane,
                    vPlane,
                    reconY,
                    reconU,
                    reconV,
                    width,
                    height,
                    chromaWidth,
                    chromaHeight,
                    macroblockOffsetX,
                    macroblockOffsetY,
                    chromaOffsetX,
                    chromaOffsetY,
                    bestY16Mode,
                    uvMode);
            } else {
                bPredModes = _macroblockModes;
                skipCoefficients = ApplyPredictionAndCheckMatchBPred(
                    yPlane,
                    uPlane,
                    vPlane,
                    reconY,
                    reconU,
                    reconV,
                    width,
                    height,
                    chromaWidth,
                    chromaHeight,
                    macroblockOffsetX,
                    macroblockOffsetY,
                    chromaOffsetX,
                    chromaOffsetY,
                    uvMode,
                    bPredModes);
            }

            if (!skipCoefficients) {
                RestorePlaneBlock(reconY, width, height, macroblockOffsetX, macroblockOffsetY, MacroblockSize, MacroblockSize, yBackup);
                RestorePlaneBlock(reconU, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, MacroblockSize / 2, MacroblockSize / 2, uBackup);
                RestorePlaneBlock(reconV, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, MacroblockSize / 2, MacroblockSize / 2, vBackup);
            }
        }

        if (enableSkip) {
            headerWriter.WriteBool(skipProbability, skipCoefficients);
        }

        if (useY16) {
            WriteKeyframeYMode(headerWriter, bestY16Mode);
            for (var i = 0; i < MacroblockSubBlockCount; i++) {
                currentSubModes[yBlockBase + i] = bestY16Mode switch { 1 => 2, 2 => 3, 3 => 1, _ => 0 };
            }
            if (skipCoefficients) {
                y2NzCurrent[mbX] = 0;
                y2Left = 0;
                for (var i = 0; i < MacroblockSubBlockCount; i++) {
                    yNzCurrent[yBlockBase + i] = 0;
                }
            } else {
                EncodeLumaMacroblockY16(
                    tokenWriter,
                    probabilities,
                    yPlane,
                    reconY,
                    width,
                    height,
                    mbX,
                    mbY,
                    bestY16Mode,
                    dequant,
                    y2NzAbove,
                    y2NzCurrent,
                    ref y2Left,
                    yNzAbove,
                    yNzCurrent);
            }
        } else {
            WriteKeyframeYMode(headerWriter, YModeBPred);
            // RFC 6386 Ã‚Â§13.3: Y2 context comes from the most recent block
            // in this row/column that has Y2, skipping B_PRED macroblocks.
            y2NzCurrent[mbX] = y2NzAbove[mbX];

            if (skipCoefficients && bPredModes != null) {
                for (var blockIndex = 0; blockIndex < MacroblockSubBlockCount; blockIndex++) {
                    var subX = blockIndex & 3;
                    var subY = blockIndex >> 2;

                    var aboveMode = BModeDcPred;
                    if (subY == 0) {
                        if (mbY > 0) {
                            aboveMode = aboveSubModes[yBlockBase + 12 + subX];
                        }
                    } else {
                        aboveMode = bPredModes[((subY - 1) * 4) + subX];
                    }

                    var leftMode = BModeDcPred;
                    if (subX == 0) {
                        if (mbX > 0) {
                            leftMode = currentSubModes[yBlockBase - MacroblockSubBlockCount + (subY * 4) + 3];
                        }
                    } else {
                        leftMode = bPredModes[(subY * 4) + subX - 1];
                    }

                    var mode = bPredModes[blockIndex];
                    WriteKeyframeBMode(headerWriter, aboveMode, leftMode, mode);
                    currentSubModes[yBlockBase + blockIndex] = mode;
                    yNzCurrent[yBlockBase + blockIndex] = 0;
                }
            } else {
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

                    var aboveMode = BModeDcPred;
                    if (subY == 0) {
                        if (mbY > 0) {
                            aboveMode = aboveSubModes[yBlockBase + 12 + subX];
                        }
                    } else {
                        aboveMode = currentSubModes[yBlockBase + ((subY - 1) * 4) + subX];
                    }

                    var leftMode = BModeDcPred;
                    if (subX == 0) {
                        if (mbX > 0) {
                            leftMode = currentSubModes[yBlockBase - MacroblockSubBlockCount + (subY * 4) + 3];
                        }
                    } else {
                        leftMode = currentSubModes[yBlockBase + (subY * 4) + subX - 1];
                    }

                    var dstX = macroblockOffsetX + (subX * BlockSize);
                    var dstY = macroblockOffsetY + (subY * BlockSize);
                    var mode = ChooseBestBlockMode(yPlane, reconY, width, height, dstX, dstY, out _);

                    WriteKeyframeBMode(headerWriter, aboveMode, leftMode, mode);
                    currentSubModes[yBlockBase + blockIndex] = mode;

                    var hasNonZero = EncodeBlock(
                        tokenWriter,
                        probabilities,
                        BlockTypeY,
                        initialContext,
                        yPlane,
                        reconY,
                        width,
                        height,
                        dstX,
                        dstY,
                        mode,
                        dequant.Y1Dc,
                        dequant.Y1Ac);

                    yNzCurrent[yBlockBase + blockIndex] = hasNonZero ? (byte)1 : (byte)0;
                }
            }
        }

        WriteKeyframeUvMode(headerWriter, uvMode);

        var chromaBlockBase = mbX * MacroblockChromaBlocks;
        if (skipCoefficients) {
            for (var i = 0; i < MacroblockChromaBlocks; i++) {
                uNzCurrent[chromaBlockBase + i] = 0;
                vNzCurrent[chromaBlockBase + i] = 0;
            }
            return;
        }
        PrefillPrediction(reconU, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, 8, uvMode);
        PrefillPrediction(reconV, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, 8, uvMode);
        for (var blockIndex = 0; blockIndex < MacroblockChromaBlocks; blockIndex++) {
            var subX = blockIndex & 1;
            var subY = blockIndex >> 1;
            var initialContext = 0;

            if (subY == 0) {
                if (mbY > 0) initialContext += uNzAbove[chromaBlockBase + 2 + subX];
            } else {
                initialContext += uNzCurrent[chromaBlockBase + ((subY - 1) * 2) + subX];
            }

            if (subX == 0) {
                if (mbX > 0) initialContext += uNzCurrent[chromaBlockBase - MacroblockChromaBlocks + (subY * 2) + 1];
            } else {
                initialContext += uNzCurrent[chromaBlockBase + (subY * 2) + subX - 1];
            }

            if (initialContext > 2) initialContext = 2;

            var dstX = chromaOffsetX + (subX * BlockSize);
            var dstY = chromaOffsetY + (subY * BlockSize);

            var hasNonZeroU = EncodeBlock(
                tokenWriter,
                probabilities,
                BlockTypeU,
                initialContext,
                uPlane,
                reconU,
                chromaWidth,
                chromaHeight,
                dstX,
                dstY,
                uvMode,
                dequant.UvDc,
                dequant.UvAc);

            uNzCurrent[chromaBlockBase + blockIndex] = hasNonZeroU ? (byte)1 : (byte)0;
        }

        for (var blockIndex = 0; blockIndex < MacroblockChromaBlocks; blockIndex++) {
            var subX = blockIndex & 1;
            var subY = blockIndex >> 1;
            var initialContext = 0;

            if (subY == 0) {
                if (mbY > 0) initialContext += vNzAbove[chromaBlockBase + 2 + subX];
            } else {
                initialContext += vNzCurrent[chromaBlockBase + ((subY - 1) * 2) + subX];
            }

            if (subX == 0) {
                if (mbX > 0) initialContext += vNzCurrent[chromaBlockBase - MacroblockChromaBlocks + (subY * 2) + 1];
            } else {
                initialContext += vNzCurrent[chromaBlockBase + (subY * 2) + subX - 1];
            }

            if (initialContext > 2) initialContext = 2;

            var dstX = chromaOffsetX + (subX * BlockSize);
            var dstY = chromaOffsetY + (subY * BlockSize);

            var hasNonZeroV = EncodeBlock(
                tokenWriter,
                probabilities,
                BlockTypeV,
                initialContext,
                vPlane,
                reconV,
                chromaWidth,
                chromaHeight,
                dstX,
                dstY,
                uvMode,
                dequant.UvDc,
                dequant.UvAc);

            vNzCurrent[chromaBlockBase + blockIndex] = hasNonZeroV ? (byte)1 : (byte)0;
        }
    }

}
