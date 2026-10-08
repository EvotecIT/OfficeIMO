// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    private int QuantizeDouble(double value, int dequant) {
        if (dequant <= 0) return 0;
        var scaled = value / dequant;
        if (scaled >= 0) return (int)Math.Floor(scaled + 0.5);
        return (int)Math.Ceiling(scaled - 0.5);
    }

    private int ClampCoefficient(int value) {
        if (value < -MaxCoefficientMagnitude) return -MaxCoefficientMagnitude;
        if (value > MaxCoefficientMagnitude) return MaxCoefficientMagnitude;
        return value;
    }

    private int QualityToBaseQIndex(int quality) {
        if (quality <= 0) return 127;
        if (quality >= 100) return 0;
        var q = quality / 100.0;
        var curve = Math.Pow(1.0 - q, 1.35);
        var index = (int)Math.Round(curve * 127);
        if (index < 0) return 0;
        if (index > 127) return 127;
        return index;
    }

    private int ComputeSkipProbability(int baseQIndex, bool segmentationEnabled) {
        var baseProbability = segmentationEnabled ? 200 : 140;
        var span = segmentationEnabled ? 40 : 30;
        var qualityFactor = baseQIndex * span / 127;
        var probability = baseProbability + qualityFactor;
        if (probability < 8) return 8;
        if (probability > 248) return 248;
        return probability;
    }

    private byte[] BuildSegmentMap(
        byte[] yPlane,
        int width,
        int height,
        int macroblockCols,
        int macroblockRows,
        int baseQIndex,
        out bool segmentationEnabled,
        out int[] quantizerDeltas,
        out int[] segmentProbabilities) {
        segmentationEnabled = false;
        quantizerDeltas = new int[SegmentCount];
        segmentProbabilities = new int[SegmentProbCount] { 128, 128, 128 };

        var macroblockCountLong = (long)macroblockCols * macroblockRows;
        if (macroblockCountLong <= 0 || macroblockCountLong > int.MaxValue) {
            return Array.Empty<byte>();
        }

        var macroblockCount = (int)macroblockCountLong;
        var segmentIds = new byte[macroblockCount];

        if (macroblockCount < 4 || baseQIndex < 12) {
            return segmentIds;
        }

        var variances = new int[macroblockCount];
        for (var mbY = 0; mbY < macroblockRows; mbY++) {
            Checkpoint();
            for (var mbX = 0; mbX < macroblockCols; mbX++) {
                var index = (mbY * macroblockCols) + mbX;
                var startX = mbX * MacroblockSize;
                var startY = mbY * MacroblockSize;

                long sum = 0;
                long sumSq = 0;
                var count = 0;

                for (var y = 0; y < MacroblockSize; y++) {
            Checkpoint();
                    var srcY = startY + y;
                    if (srcY >= height) break;
                    var rowOffset = srcY * width;
                    for (var x = 0; x < MacroblockSize; x++) {
                        var srcX = startX + x;
                        if (srcX >= width) break;
                        var sample = yPlane[rowOffset + srcX];
                        sum += sample;
                        sumSq += (long)sample * sample;
                        count++;
                    }
                }

                if (count <= 0) {
                    variances[index] = 0;
                    continue;
                }

                var numerator = (sumSq * count) - (sum * sum);
                if (numerator < 0) numerator = 0;
                variances[index] = (int)(numerator / (count * (long)count));
            }
        }

        var sorted = (int[])variances.Clone();
        Array.Sort(sorted);
        var p25 = sorted[(macroblockCount * 1) / 4];
        var p50 = sorted[(macroblockCount * 2) / 4];
        var p75 = sorted[(macroblockCount * 3) / 4];

        var spread = p75 - p25;
        var minSpread = Math.Max(8, baseQIndex / 2);
        if (spread < minSpread) {
            return segmentIds;
        }

        var counts = new int[SegmentCount];
        for (var i = 0; i < variances.Length; i++) {
            var variance = variances[i];
            var segmentId = variance <= p25 ? 0
                : variance <= p50 ? 1
                : variance <= p75 ? 2
                : 3;
            segmentIds[i] = (byte)segmentId;
            counts[segmentId]++;
        }

        var usedSegments = 0;
        for (var i = 0; i < counts.Length; i++) {
            if (counts[i] > 0) usedSegments++;
        }

        if (usedSegments < 2) {
            return segmentIds;
        }

        var step = baseQIndex / 16 + spread / 32;
        if (step < 2) step = 2;
        if (step > 16) step = 16;

        quantizerDeltas[0] = ClampSegmentDelta(step * 2);
        quantizerDeltas[1] = ClampSegmentDelta(step);
        quantizerDeltas[2] = 0;
        quantizerDeltas[3] = ClampSegmentDelta(-step);

        segmentProbabilities = BuildSegmentProbabilities(counts);
        segmentationEnabled = true;
        return segmentIds;
    }

    private int[] BuildSegmentProbabilities(int[] counts) {
        var probabilities = new int[SegmentProbCount] { 128, 128, 128 };
        if (counts == null || counts.Length < SegmentCount) return probabilities;

        var total = 0;
        for (var i = 0; i < SegmentCount; i++) {
            total += counts[i];
        }

        if (total <= 0) return probabilities;

        probabilities[0] = ComputeProbability(counts[0] + counts[1], total);
        probabilities[1] = ComputeProbability(counts[0], counts[0] + counts[1]);
        probabilities[2] = ComputeProbability(counts[2], counts[2] + counts[3]);
        return probabilities;
    }

    private int ComputeProbability(int countFalse, int total) {
        if (total <= 0) return 128;
        var numerator = (long)countFalse * 255 + (total / 2);
        var prob = (int)(numerator / total);
        if (prob < 0) return 0;
        if (prob > 255) return 255;
        return prob;
    }

    private DequantFactors[] BuildDequantFactors(int baseQIndex, int[] quantizerDeltas, bool segmentationEnabled) {
        var factors = new DequantFactors[SegmentCount];
        if (!segmentationEnabled) {
            var baseFactors = BuildDequantFactors(baseQIndex);
            for (var i = 0; i < SegmentCount; i++) {
                factors[i] = baseFactors;
            }
            return factors;
        }

        for (var i = 0; i < SegmentCount; i++) {
            var delta = (quantizerDeltas != null && i < quantizerDeltas.Length) ? quantizerDeltas[i] : 0;
            var qIndex = ClampQIndex(baseQIndex + delta);
            factors[i] = BuildDequantFactors(qIndex);
        }

        return factors;
    }

    private DequantFactors BuildDequantFactors(int baseQIndex) {
        var q = ClampQIndex(baseQIndex);
        var y1Dc = GetDcQuant(q);
        var y1Ac = GetAcQuant(q);
        var y2Dc = GetDcQuant(q) * 2;
        var y2Ac = (GetAcQuant(q) * 155) / 100;
        if (y2Ac < 8) y2Ac = 8;
        var uvDc = GetDcQuant(q);
        if (uvDc > 132) uvDc = 132;
        var uvAc = GetAcQuant(q);
        return new DequantFactors(y1Dc, y1Ac, y2Dc, y2Ac, uvDc, uvAc);
    }

    private int GetDcQuant(int qIndex) {
        var clamped = ClampQIndex(qIndex);
        if ((uint)clamped >= (uint)OfficeVp8Tables.DcQlookup.Length) return 0;
        return OfficeVp8Tables.DcQlookup[clamped];
    }

    private int GetAcQuant(int qIndex) {
        var clamped = ClampQIndex(qIndex);
        if ((uint)clamped >= (uint)OfficeVp8Tables.AcQlookup.Length) return 0;
        return OfficeVp8Tables.AcQlookup[clamped];
    }

    private int ClampQIndex(int qIndex) {
        if (qIndex < 0) return 0;
        if (qIndex > 127) return 127;
        return qIndex;
    }

    private int ClampSegmentDelta(int delta) {
        if (delta < -127) return -127;
        if (delta > 127) return 127;
        return delta;
    }

    private int GetCoeffIndex(int blockType, int band, int prev, int node) {
        return (((blockType * CoeffBands) + band) * CoeffPrevContexts + prev) * CoeffEntropyNodes + node;
    }

    private int GetMacroblockDimension(int pixels) {
        if (pixels <= 0) return 0;
        return (pixels + MacroblockSize - 1) / MacroblockSize;
    }

    private byte ClampToByte(int value) {
        if (value < byte.MinValue) return byte.MinValue;
        if (value > byte.MaxValue) return byte.MaxValue;
        return (byte)value;
    }

}
