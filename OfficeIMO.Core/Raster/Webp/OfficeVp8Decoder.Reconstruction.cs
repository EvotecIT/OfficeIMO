// Adapted from CodeGlyphX, commit fc25e2fcf795d9c9a09b88708c47bdfeed5c446d.
// Copyright CodeGlyphX contributors. Apache-2.0; see THIRD-PARTY-NOTICES.md.
// OfficeIMO adaptation removes diagnostic scaffolding and adds bounded cancellation/resource handling.
using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeVp8Decoder {
    private static bool TryDecodeMacroblock(
        OfficeVp8BoolDecoder decoder,
        OfficeVp8CoefficientProbabilities probabilities,
        OfficeVp8MacroblockHeader macroblock,
        DequantFactors dequant,
        int macroblockX,
        int macroblockY,
        int width,
        int height,
        byte[] yPlane,
        byte[] uPlane,
        byte[] vPlane,
        int chromaWidth,
        int chromaHeight,
        byte[] y2NzAbove,
        byte[] y2NzCurrent,
        ref byte y2Left,
        byte[] yNzAbove,
        byte[] yNzCurrent,
        byte[] uNzAbove,
        byte[] uNzCurrent,
        byte[] vNzAbove,
        byte[] vNzCurrent,
        OfficeVp8PredictionScratch predictionScratch,
        out bool hasCoefficients) {
        hasCoefficients = false;
        var mbX = macroblockX;
        var mbY = macroblockY;
        var macroblockOffsetX = mbX * MacroblockSize;
        var macroblockOffsetY = mbY * MacroblockSize;
        var chromaOffsetX = mbX * MacroblockChromaWidth;
        var chromaOffsetY = mbY * MacroblockChromaHeight;

        var skipCoefficients = macroblock.SkipCoefficients;
        var hasY2 = !macroblock.Is4x4;
        if (hasY2) PredictDecodedBlock(yPlane, width, height, macroblockOffsetX, macroblockOffsetY, 16, macroblock.YMode, predictionScratch);
        PredictDecodedBlock(uPlane, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, 8, macroblock.UvMode, predictionScratch);
        PredictDecodedBlock(vPlane, chromaWidth, chromaHeight, chromaOffsetX, chromaOffsetY, 8, macroblock.UvMode, predictionScratch);


        var y2Dc = ZeroCoefficients;
        if (hasY2) {
            var initialContext = 0;
            if (mbY > 0) initialContext += y2NzAbove[mbX];
            if (mbX > 0) initialContext += y2Left;
            if (initialContext > 2) initialContext = 2;

            var y2NonZero = false;
            var y2Coeffs = ZeroCoefficients;
            if (!skipCoefficients) {
                if (!TryDecodeBlockCoefficients(
                    decoder,
                    probabilities,
                    OfficeVp8Tables.LumaSecondOrder,
                    initialContext,
                    dequant.Y2Dc,
                    dequant.Y2Ac,
                    out y2Coeffs,
                    out y2NonZero)) {
                    return false;
                }
            }

            y2Left = y2NonZero ? (byte)1 : (byte)0;
            y2NzCurrent[mbX] = y2Left;
            if (y2NonZero) {
                y2Dc = OfficeVp8Transform.InverseWalshTransform4x4(y2Coeffs);
            }

            hasCoefficients |= y2NonZero;
        } else {
            // RFC 6386 §13.3: Y2 context comes from the most recent block
            // in this row/column that has Y2, skipping B_PRED macroblocks.
            y2NzCurrent[mbX] = y2NzAbove[mbX];
        }

        var yBlockBase = mbX * MacroblockSubBlockCount;
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

            var hasNonZero = false;
            var coefficients = ZeroCoefficients;
            if (!skipCoefficients) {
                if (!TryDecodeBlockCoefficients(
                    decoder,
                    probabilities,
                    hasY2 ? OfficeVp8Tables.LumaAc : OfficeVp8Tables.LumaWithDc,
                    initialContext,
                    dequant.Y1Dc,
                    dequant.Y1Ac,
                    out coefficients,
                    out hasNonZero)) {
                    return false;
                }
            }

            yNzCurrent[yBlockBase + blockIndex] = hasNonZero ? (byte)1 : (byte)0;
            hasCoefficients |= hasNonZero;
            var mode = macroblock.Is4x4
                ? GetSubblockMode(macroblock.SubblockModes, blockIndex)
                : macroblock.YMode;

            var dstX = macroblockOffsetX + (subX * BlockSize);
            var dstY = macroblockOffsetY + (subY * BlockSize);
            var dcOverride = hasY2 && y2Dc.Length > blockIndex ? y2Dc[blockIndex] : 0;

            ApplyDecodedResidual(
                yPlane,
                width,
                height,
                dstX,
                dstY,
                coefficients,
                macroblock.Is4x4,
                mode,
                hasY2,
                dcOverride, predictionScratch);
        }

        var chromaBlockBase = mbX * MacroblockChromaBlocks;
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

            var hasNonZero = false;
            var coefficients = ZeroCoefficients;
            if (!skipCoefficients) {
                if (!TryDecodeBlockCoefficients(
                    decoder,
                    probabilities,
                    OfficeVp8Tables.Chroma,
                    initialContext,
                    dequant.UvDc,
                    dequant.UvAc,
                    out coefficients,
                    out hasNonZero)) {
                    return false;
                }
            }

            uNzCurrent[chromaBlockBase + blockIndex] = hasNonZero ? (byte)1 : (byte)0;
            hasCoefficients |= hasNonZero;

            var dstX = chromaOffsetX + (subX * BlockSize);
            var dstY = chromaOffsetY + (subY * BlockSize);
            ApplyDecodedResidual(
                uPlane,
                chromaWidth,
                chromaHeight,
                dstX,
                dstY,
                coefficients,
                is4x4: false,
                macroblock.UvMode,
                overrideDc: false,
                dcValue: 0, predictionScratch);
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

            var hasNonZero = false;
            var coefficients = ZeroCoefficients;
            if (!skipCoefficients) {
                if (!TryDecodeBlockCoefficients(
                    decoder,
                    probabilities,
                    OfficeVp8Tables.Chroma,
                    initialContext,
                    dequant.UvDc,
                    dequant.UvAc,
                    out coefficients,
                    out hasNonZero)) {
                    return false;
                }
            }

            vNzCurrent[chromaBlockBase + blockIndex] = hasNonZero ? (byte)1 : (byte)0;
            hasCoefficients |= hasNonZero;

            var dstX = chromaOffsetX + (subX * BlockSize);
            var dstY = chromaOffsetY + (subY * BlockSize);
            ApplyDecodedResidual(
                vPlane,
                chromaWidth,
                chromaHeight,
                dstX,
                dstY,
                coefficients,
                is4x4: false,
                macroblock.UvMode,
                overrideDc: false,
                dcValue: 0, predictionScratch);
        }

        return true;
    }

    private static DequantFactors[] BuildDequantFactors(OfficeVp8Quantization quantization, OfficeVp8Segmentation segmentation) {
        var factors = new DequantFactors[SegmentCount];
        var segmentCount = segmentation.Enabled ? SegmentCount : 1;

        for (var i = 0; i < segmentCount; i++) {
            var q = quantization.BaseQIndex;
            if (segmentation.Enabled) {
                var delta = (i < segmentation.QuantizerDeltas.Length) ? segmentation.QuantizerDeltas[i] : 0;
                q = segmentation.AbsoluteDeltas ? delta : q + delta;
            }

            var y1Dc = GetDcQuant(q + GetDelta(quantization, 0));
            var y1Ac = GetAcQuant(q);
            var y2Dc = GetDcQuant(q + GetDelta(quantization, 1)) * 2;
            var y2Ac = (GetAcQuant(q + GetDelta(quantization, 2)) * 155) / 100;
            if (y2Ac < 8) y2Ac = 8;
            var uvDc = GetDcQuant(q + GetDelta(quantization, 3));
            if (uvDc > 132) uvDc = 132;
            var uvAc = GetAcQuant(q + GetDelta(quantization, 4));

            factors[i] = new DequantFactors(y1Dc, y1Ac, y2Dc, y2Ac, uvDc, uvAc);
        }

        if (!segmentation.Enabled) {
            for (var i = 1; i < factors.Length; i++) {
                factors[i] = factors[0];
            }
        }

        return factors;
    }

    private static int GetDelta(OfficeVp8Quantization quantization, int index) {
        if (quantization.Deltas is null || (uint)index >= (uint)quantization.Deltas.Length) return 0;
        return quantization.Deltas[index];
    }

    private static int GetDcQuant(int qIndex) {
        var clamped = ClampQIndex(qIndex);
        if ((uint)clamped >= (uint)OfficeVp8Tables.DcQlookup.Length) return 0;
        return OfficeVp8Tables.DcQlookup[clamped];
    }

    private static int GetAcQuant(int qIndex) {
        var clamped = ClampQIndex(qIndex);
        if ((uint)clamped >= (uint)OfficeVp8Tables.AcQlookup.Length) return 0;
        return OfficeVp8Tables.AcQlookup[clamped];
    }

    private static void SwapRowContexts(ref byte[] above, ref byte[] current) {
        var temp = above;
        above = current;
        current = temp;
        Array.Clear(current, 0, current.Length);
    }

    private static void SwapRowModes(ref int[] above, ref int[] current) {
        var temp = above;
        above = current;
        current = temp;
        Array.Clear(current, 0, current.Length);
    }

    private static byte[] ConvertPlanesToRgba(
        int width,
        int height,
        int lumaStride,
        byte[] yPlane,
        byte[] uPlane,
        byte[] vPlane,
        int chromaWidth,
        int chromaHeight,
        CancellationToken cancellationToken) {
        var rgba = new byte[checked(width * height * 4)];

        for (var y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (var x = 0; x < width; x++) {
                var ySample = yPlane[(y * lumaStride) + x];
                var uSample = InterpolateDecodedChroma(uPlane, chromaWidth, (width + 1) >> 1, (height + 1) >> 1, x, y);
                var vSample = InterpolateDecodedChroma(vPlane, chromaWidth, (width + 1) >> 1, (height + 1) >> 1, x, y);

                var (r, g, b) = ConvertYuvToRgb(ySample, uSample, vSample);
                var dst = (y * width + x) * 4;
                rgba[dst + 0] = r;
                rgba[dst + 1] = g;
                rgba[dst + 2] = b;
                rgba[dst + 3] = 255;
            }
        }

        return rgba;
    }

    private static OfficeVp8TokenInfo ClassifyToken(int tokenCode, int band, int prevContextBefore) {
        var hasMore = tokenCode != 0;
        var isNonZero = tokenCode >= 2;
        var prevContextAfter = tokenCode switch {
            0 or 1 => 0,
            2 => 1,
            _ => 2,
        };

        return new OfficeVp8TokenInfo(tokenCode, band, prevContextBefore, prevContextAfter, hasMore, isNonZero);
    }

    private static int MapZigZagToNaturalIndex(int zigZagIndex) {
        if ((uint)zigZagIndex >= ZigZagToNaturalOrder.Length) return 0;
        return ZigZagToNaturalOrder[zigZagIndex];
    }

    private static byte ClampToByte(int value) {
        if (value < byte.MinValue) return byte.MinValue;
        if (value > byte.MaxValue) return byte.MaxValue;
        return (byte)value;
    }

    private static int GetSubblockMode(int[] modes, int index) {
        if (modes is null || modes.Length == 0) return 0;
        if ((uint)index >= (uint)modes.Length) return 0;
        return modes[index];
    }

    private static int ClampQIndex(int qIndex) {
        if (qIndex < 0) return 0;
        if (qIndex > 127) return 127;
        return qIndex;
    }

    private static int GetMacroblockDimension(int pixels) {
        if (pixels <= 0) return 0;
        return (pixels + MacroblockSize - 1) / MacroblockSize;
    }

    private static int GetTokenPartitionForRow(int macroblockRow, int partitionCount) {
        if (partitionCount <= 0) return -1;
        if (macroblockRow < 0) macroblockRow = 0;
        // VP8 uses power-of-two partition counts; modulo selects the coefficient partition for each macroblock row.
        return macroblockRow % partitionCount;
    }

    private static (byte R, byte G, byte B) ConvertYuvToRgb(byte y, byte u, byte v) {
        var yy = y - 16;
        var uu = u - 128;
        var vv = v - 128;

        var r = (76284 * yy + 104595 * vv + 32768) >> 16;
        var g = (76284 * yy - 25624 * uu - 53281 * vv + 32768) >> 16;
        var b = (76284 * yy + 132251 * uu + 32768) >> 16;

        return (ClampToByte(r), ClampToByte(g), ClampToByte(b));
    }
}
