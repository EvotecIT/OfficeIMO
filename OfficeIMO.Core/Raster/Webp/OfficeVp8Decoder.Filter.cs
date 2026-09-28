// Adapted from CodeGlyphX, commit fc25e2fcf795d9c9a09b88708c47bdfeed5c446d.
// Copyright CodeGlyphX contributors. Apache-2.0; see THIRD-PARTY-NOTICES.md.
// OfficeIMO adaptation removes diagnostic scaffolding and adds bounded cancellation/resource handling.
using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeVp8Decoder {
    private static void ApplyLoopFilter(
        OfficeVp8LoopFilter loopFilter,
        OfficeVp8Segmentation segmentation,
        OfficeVp8MacroblockHeader[] macroblocks,
        bool[] macroblockHasCoefficients,
        int width,
        int height,
        byte[] yPlane,
        byte[] uPlane,
        byte[] vPlane,
        int chromaWidth,
        int chromaHeight,
        bool isKeyframe, CancellationToken cancellationToken) {
        if (macroblocks.Length == 0) return;
        if (loopFilter.Level <= 0 && (!segmentation.Enabled || segmentation.FilterDeltas.Length == 0)) {
            return;
        }

        var macroblockCols = GetMacroblockDimension(width);
        var macroblockRows = GetMacroblockDimension(height);
        if (macroblockCols <= 0 || macroblockRows <= 0) return;

        var useSimple = loopFilter.FilterType != 0;
        var strideY = width;
        var strideUv = chromaWidth;

        for (var row = 0; row < macroblockRows; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (var col = 0; col < macroblockCols; col++) {
                var index = (row * macroblockCols) + col;
                if ((uint)index >= (uint)macroblocks.Length) continue;
                var macroblock = macroblocks[index];
                var hasCoefficients = (uint)index < (uint)macroblockHasCoefficients.Length && macroblockHasCoefficients[index];
                var filterSubblocks = hasCoefficients || macroblock.Is4x4;

                if (!TryCalculateLoopFilterParameters(
                    loopFilter,
                    segmentation,
                    macroblock,
                    isKeyframe,
                    out var edgeLimit,
                    out var interiorLimit,
                    out var hevThreshold)) {
                    continue;
                }

                if (edgeLimit <= 0) continue;

                var yX = col * MacroblockSize;
                var yY = row * MacroblockSize;
                var uvX = col * MacroblockChromaWidth;
                var uvY = row * MacroblockChromaHeight;

                if (useSimple) {
                    var mbLimit = ((edgeLimit + 2) * 2) + interiorLimit;
                    var bLimit = (edgeLimit * 2) + interiorLimit;

                    if (col > 0) {
                        FilterVEdgeSimple(yPlane, width, height, strideY, yX, yY, mbLimit);
                    }

                    if (filterSubblocks) {
                        FilterVEdgeSimple(yPlane, width, height, strideY, yX + 4, yY, bLimit);
                        FilterVEdgeSimple(yPlane, width, height, strideY, yX + 8, yY, bLimit);
                        FilterVEdgeSimple(yPlane, width, height, strideY, yX + 12, yY, bLimit);
                    }

                    if (row > 0) {
                        FilterHEdgeSimple(yPlane, width, height, strideY, yX, yY, mbLimit);
                    }

                    if (filterSubblocks) {
                        FilterHEdgeSimple(yPlane, width, height, strideY, yX, yY + 4, bLimit);
                        FilterHEdgeSimple(yPlane, width, height, strideY, yX, yY + 8, bLimit);
                        FilterHEdgeSimple(yPlane, width, height, strideY, yX, yY + 12, bLimit);
                    }
                } else {
                    if (col > 0) {
                        FilterMbVEdge(yPlane, width, height, strideY, yX, yY, edgeLimit + 2, interiorLimit, hevThreshold, size: 2);
                        FilterMbVEdge(uPlane, chromaWidth, chromaHeight, strideUv, uvX, uvY, edgeLimit + 2, interiorLimit, hevThreshold, size: 1);
                        FilterMbVEdge(vPlane, chromaWidth, chromaHeight, strideUv, uvX, uvY, edgeLimit + 2, interiorLimit, hevThreshold, size: 1);
                    }

                    if (filterSubblocks) {
                        FilterSubblockVEdge(yPlane, width, height, strideY, yX + 4, yY, edgeLimit, interiorLimit, hevThreshold, size: 2);
                        FilterSubblockVEdge(yPlane, width, height, strideY, yX + 8, yY, edgeLimit, interiorLimit, hevThreshold, size: 2);
                        FilterSubblockVEdge(yPlane, width, height, strideY, yX + 12, yY, edgeLimit, interiorLimit, hevThreshold, size: 2);
                        FilterSubblockVEdge(uPlane, chromaWidth, chromaHeight, strideUv, uvX + 4, uvY, edgeLimit, interiorLimit, hevThreshold, size: 1);
                        FilterSubblockVEdge(vPlane, chromaWidth, chromaHeight, strideUv, uvX + 4, uvY, edgeLimit, interiorLimit, hevThreshold, size: 1);
                    }

                    if (row > 0) {
                        FilterMbHEdge(yPlane, width, height, strideY, yX, yY, edgeLimit + 2, interiorLimit, hevThreshold, size: 2);
                        FilterMbHEdge(uPlane, chromaWidth, chromaHeight, strideUv, uvX, uvY, edgeLimit + 2, interiorLimit, hevThreshold, size: 1);
                        FilterMbHEdge(vPlane, chromaWidth, chromaHeight, strideUv, uvX, uvY, edgeLimit + 2, interiorLimit, hevThreshold, size: 1);
                    }

                    if (filterSubblocks) {
                        FilterSubblockHEdge(yPlane, width, height, strideY, yX, yY + 4, edgeLimit, interiorLimit, hevThreshold, size: 2);
                        FilterSubblockHEdge(yPlane, width, height, strideY, yX, yY + 8, edgeLimit, interiorLimit, hevThreshold, size: 2);
                        FilterSubblockHEdge(yPlane, width, height, strideY, yX, yY + 12, edgeLimit, interiorLimit, hevThreshold, size: 2);
                        FilterSubblockHEdge(uPlane, chromaWidth, chromaHeight, strideUv, uvX, uvY + 4, edgeLimit, interiorLimit, hevThreshold, size: 1);
                        FilterSubblockHEdge(vPlane, chromaWidth, chromaHeight, strideUv, uvX, uvY + 4, edgeLimit, interiorLimit, hevThreshold, size: 1);
                    }
                }
            }
        }
    }

    private static bool TryCalculateLoopFilterParameters(
        OfficeVp8LoopFilter loopFilter,
        OfficeVp8Segmentation segmentation,
        OfficeVp8MacroblockHeader macroblock,
        bool isKeyframe,
        out int edgeLimit,
        out int interiorLimit,
        out int hevThreshold) {
        edgeLimit = 0;
        interiorLimit = 0;
        hevThreshold = 0;

        var filterLevel = loopFilter.Level;
        if (segmentation.Enabled && segmentation.FilterDeltas.Length > 0) {
            var segmentId = macroblock.SegmentId;
            if ((uint)segmentId >= SegmentCount) segmentId = 0;
            var delta = segmentation.FilterDeltas[segmentId];
            filterLevel = segmentation.AbsoluteDeltas ? delta : filterLevel + delta;
        }

        if (loopFilter.DeltaEnabled) {
            var refDelta = loopFilter.RefDeltas.Length > 0 ? loopFilter.RefDeltas[0] : 0;
            filterLevel += refDelta;
            if (macroblock.Is4x4 && loopFilter.ModeDeltas.Length > 0) {
                filterLevel += loopFilter.ModeDeltas[0];
            }
        }

        // Combine all signed deltas before applying the VP8 six-bit filter bound.
        filterLevel = Math.Max(0, Math.Min(63, filterLevel));
        if (filterLevel == 0) return false;

        interiorLimit = filterLevel;
        if (loopFilter.Sharpness > 0) {
            interiorLimit >>= loopFilter.Sharpness > 4 ? 2 : 1;
            var clamp = 9 - loopFilter.Sharpness;
            if (interiorLimit > clamp) interiorLimit = clamp;
        }
        if (interiorLimit < 1) interiorLimit = 1;

        hevThreshold = filterLevel >= 15 ? 1 : 0;
        if (filterLevel >= 40) hevThreshold++;
        if (!isKeyframe && filterLevel >= 20) hevThreshold++;

        edgeLimit = filterLevel;
        return true;
    }

    private static void FilterMbVEdge(
        byte[] plane,
        int width,
        int height,
        int stride,
        int x,
        int y,
        int edgeLimit,
        int interiorLimit,
        int hevThreshold,
        int size) {
        if (x < 4 || x + 3 >= width) return;
        var rows = Math.Min(8 * size, height - y);
        if (rows <= 0) return;

        for (var i = 0; i < rows; i++) {
            var py = y + i;
            var index = (py * stride) + x;
            if (NormalThreshold(plane, index, step: 1, edgeLimit, interiorLimit)) {
                if (HighEdgeVariance(plane, index, step: 1, hevThreshold)) {
                    FilterCommon(plane, index, step: 1, useOuterTaps: true);
                } else {
                    FilterMbEdge(plane, index, step: 1);
                }
            }
        }
    }

    private static void FilterSubblockVEdge(
        byte[] plane,
        int width,
        int height,
        int stride,
        int x,
        int y,
        int edgeLimit,
        int interiorLimit,
        int hevThreshold,
        int size) {
        if (x < 4 || x + 3 >= width) return;
        var rows = Math.Min(8 * size, height - y);
        if (rows <= 0) return;

        for (var i = 0; i < rows; i++) {
            var py = y + i;
            var index = (py * stride) + x;
            if (NormalThreshold(plane, index, step: 1, edgeLimit, interiorLimit)) {
                var useOuter = HighEdgeVariance(plane, index, step: 1, hevThreshold);
                FilterCommon(plane, index, step: 1, useOuterTaps: useOuter);
            }
        }
    }

    private static void FilterMbHEdge(
        byte[] plane,
        int width,
        int height,
        int stride,
        int x,
        int y,
        int edgeLimit,
        int interiorLimit,
        int hevThreshold,
        int size) {
        if (y < 4 || y + 3 >= height) return;
        var cols = Math.Min(8 * size, width - x);
        if (cols <= 0) return;

        for (var i = 0; i < cols; i++) {
            var px = x + i;
            var index = (y * stride) + px;
            if (NormalThreshold(plane, index, step: stride, edgeLimit, interiorLimit)) {
                if (HighEdgeVariance(plane, index, step: stride, hevThreshold)) {
                    FilterCommon(plane, index, step: stride, useOuterTaps: true);
                } else {
                    FilterMbEdge(plane, index, step: stride);
                }
            }
        }
    }

    private static void FilterSubblockHEdge(
        byte[] plane,
        int width,
        int height,
        int stride,
        int x,
        int y,
        int edgeLimit,
        int interiorLimit,
        int hevThreshold,
        int size) {
        if (y < 4 || y + 3 >= height) return;
        var cols = Math.Min(8 * size, width - x);
        if (cols <= 0) return;

        for (var i = 0; i < cols; i++) {
            var px = x + i;
            var index = (y * stride) + px;
            if (NormalThreshold(plane, index, step: stride, edgeLimit, interiorLimit)) {
                var useOuter = HighEdgeVariance(plane, index, step: stride, hevThreshold);
                FilterCommon(plane, index, step: stride, useOuterTaps: useOuter);
            }
        }
    }

    private static void FilterVEdgeSimple(
        byte[] plane,
        int width,
        int height,
        int stride,
        int x,
        int y,
        int filterLimit) {
        if (x < 2 || x + 1 >= width) return;
        var rows = Math.Min(16, height - y);
        if (rows <= 0) return;

        for (var i = 0; i < rows; i++) {
            var py = y + i;
            var index = (py * stride) + x;
            if (SimpleThreshold(plane, index, step: 1, filterLimit)) {
                FilterCommon(plane, index, step: 1, useOuterTaps: true);
            }
        }
    }

    private static void FilterHEdgeSimple(
        byte[] plane,
        int width,
        int height,
        int stride,
        int x,
        int y,
        int filterLimit) {
        if (y < 2 || y + 1 >= height) return;
        var cols = Math.Min(16, width - x);
        if (cols <= 0) return;

        for (var i = 0; i < cols; i++) {
            var px = x + i;
            var index = (y * stride) + px;
            if (SimpleThreshold(plane, index, step: stride, filterLimit)) {
                FilterCommon(plane, index, step: stride, useOuterTaps: true);
            }
        }
    }

    private static bool SimpleThreshold(byte[] plane, int index, int step, int filterLimit) {
        var p0 = plane[index - step];
        var q0 = plane[index];
        var p1 = plane[index - (2 * step)];
        var q1 = plane[index + step];
        var value = (Abs(p0 - q0) * 2) + (Abs(p1 - q1) >> 1);
        return value <= filterLimit;
    }

    private static bool NormalThreshold(byte[] plane, int index, int step, int edgeLimit, int interiorLimit) {
        var limit = (2 * edgeLimit) + interiorLimit;
        if (!SimpleThreshold(plane, index, step, limit)) return false;

        var p0 = plane[index - step];
        var q0 = plane[index];
        var p1 = plane[index - (2 * step)];
        var q1 = plane[index + step];
        var p2 = plane[index - (3 * step)];
        var q2 = plane[index + (2 * step)];
        var p3 = plane[index - (4 * step)];
        var q3 = plane[index + (3 * step)];

        return Abs(p3 - p2) <= interiorLimit
            && Abs(p2 - p1) <= interiorLimit
            && Abs(p1 - p0) <= interiorLimit
            && Abs(q3 - q2) <= interiorLimit
            && Abs(q2 - q1) <= interiorLimit
            && Abs(q1 - q0) <= interiorLimit;
    }

    private static bool HighEdgeVariance(byte[] plane, int index, int step, int hevThreshold) {
        var p0 = plane[index - step];
        var q0 = plane[index];
        var p1 = plane[index - (2 * step)];
        var q1 = plane[index + step];
        return Abs(p1 - p0) > hevThreshold || Abs(q1 - q0) > hevThreshold;
    }

    private static void FilterCommon(byte[] plane, int index, int step, bool useOuterTaps) {
        var p1 = plane[index - (2 * step)];
        var p0 = plane[index - step];
        var q0 = plane[index];
        var q1 = plane[index + step];

        var a = 3 * (q0 - p0);
        if (useOuterTaps) {
            a += SaturateInt8(p1 - q1);
        }
        a = SaturateInt8(a);

        var f1 = ((a + 4 > 127) ? 127 : a + 4) >> 3;
        var f2 = ((a + 3 > 127) ? 127 : a + 3) >> 3;

        plane[index - step] = SaturateUInt8(p0 + f2);
        plane[index] = SaturateUInt8(q0 - f1);

        if (!useOuterTaps) {
            var adjust = (f1 + 1) >> 1;
            plane[index - (2 * step)] = SaturateUInt8(p1 + adjust);
            plane[index + step] = SaturateUInt8(q1 - adjust);
        }
    }

    private static void FilterMbEdge(byte[] plane, int index, int step) {
        var p2 = plane[index - (3 * step)];
        var p1 = plane[index - (2 * step)];
        var p0 = plane[index - step];
        var q0 = plane[index];
        var q1 = plane[index + step];
        var q2 = plane[index + (2 * step)];

        var w = SaturateInt8(SaturateInt8(p1 - q1) + (3 * (q0 - p0)));

        var a = (27 * w + 63) >> 7;
        plane[index - step] = SaturateUInt8(p0 + a);
        plane[index] = SaturateUInt8(q0 - a);

        a = (18 * w + 63) >> 7;
        plane[index - (2 * step)] = SaturateUInt8(p1 + a);
        plane[index + step] = SaturateUInt8(q1 - a);

        a = (9 * w + 63) >> 7;
        plane[index - (3 * step)] = SaturateUInt8(p2 + a);
        plane[index + (2 * step)] = SaturateUInt8(q2 - a);
    }

    private static int SaturateInt8(int value) {
        if (value < -128) return -128;
        if (value > 127) return 127;
        return value;
    }

    private static byte SaturateUInt8(int value) {
        if (value < 0) return 0;
        if (value > 255) return 255;
        return (byte)value;
    }

    private static bool TryReadLoopFilter(OfficeVp8BoolDecoder decoder, out OfficeVp8LoopFilter loopFilter) {
        loopFilter = default;
        if (!decoder.TryReadBool(128, out var filterTypeBit)) return false;
        if (!decoder.TryReadLiteral(6, out var level)) return false;
        if (!decoder.TryReadLiteral(3, out var sharpness)) return false;
        if (level < 0 || level > 63) return false;
        if (sharpness < 0 || sharpness > 7) return false;

        if (!decoder.TryReadBool(128, out var deltaEnabledBit)) return false;
        var deltaEnabled = deltaEnabledBit;
        var deltaUpdate = false;
        var refDeltas = new int[LoopFilterDeltaCount];
        var refDeltasUpdated = new bool[LoopFilterDeltaCount];
        var modeDeltas = new int[LoopFilterDeltaCount];
        var modeDeltasUpdated = new bool[LoopFilterDeltaCount];

        if (deltaEnabled) {
            if (!decoder.TryReadBool(128, out var deltaUpdateBit)) return false;
            deltaUpdate = deltaUpdateBit;

            if (deltaUpdate) {
                for (var i = 0; i < LoopFilterDeltaCount; i++) {
                    if (!decoder.TryReadBool(128, out var updateRefDeltaBit)) return false;
                    if (updateRefDeltaBit) {
                        if (!decoder.TryReadSignedLiteral(6, out refDeltas[i])) return false;
                        refDeltasUpdated[i] = true;
                    }
                }

                for (var i = 0; i < LoopFilterDeltaCount; i++) {
                    if (!decoder.TryReadBool(128, out var updateModeDeltaBit)) return false;
                    if (updateModeDeltaBit) {
                        if (!decoder.TryReadSignedLiteral(6, out modeDeltas[i])) return false;
                        modeDeltasUpdated[i] = true;
                    }
                }
            }
        }

        loopFilter = new OfficeVp8LoopFilter(
            filterTypeBit ? 1 : 0,
            level,
            sharpness,
            deltaEnabled,
            deltaUpdate,
            refDeltas,
            refDeltasUpdated,
            modeDeltas,
            modeDeltasUpdated);
        return true;
    }

    private static int Abs(int value) => value < 0 ? -value : value;
}
