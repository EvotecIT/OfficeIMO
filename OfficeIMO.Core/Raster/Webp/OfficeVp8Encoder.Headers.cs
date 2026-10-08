// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeVp8Encoder {
    private void WriteControlHeader(OfficeVp8BoolEncoder writer) {
        writer.WriteBool(128, false); // color space
        writer.WriteBool(128, false); // clamp type
    }

    private void WriteSegmentation(
        OfficeVp8BoolEncoder writer,
        bool enabled,
        int[] quantizerDeltas,
        int[] segmentProbabilities) {
        writer.WriteBool(128, enabled);
        if (!enabled) return;

        writer.WriteBool(128, true); // update map
        writer.WriteBool(128, true); // update data
        writer.WriteBool(128, false); // absolute deltas disabled

        for (var i = 0; i < SegmentCount; i++) {
            writer.WriteBool(128, true); // update quantizer for segment
            var delta = (quantizerDeltas != null && i < quantizerDeltas.Length) ? quantizerDeltas[i] : 0;
            WriteSignedLiteral(writer, ClampSegmentDelta(delta), 7);
        }

        for (var i = 0; i < SegmentCount; i++) {
            writer.WriteBool(128, false); // no filter delta updates
        }

        for (var i = 0; i < SegmentProbCount; i++) {
            writer.WriteBool(128, true);
            var prob = (segmentProbabilities != null && i < segmentProbabilities.Length)
                ? NormalizeSegmentProbability(segmentProbabilities[i])
                : 128;
            writer.WriteLiteral(prob, 8);
        }
    }

    private void WriteSegmentId(OfficeVp8BoolEncoder writer, int[] probabilities, int segmentId) {
        if ((uint)segmentId >= SegmentCount) segmentId = 0;
        writer.WriteBool(probabilities[0], segmentId >= 2);
        writer.WriteBool(probabilities[segmentId >= 2 ? 2 : 1], (segmentId & 1) != 0);
    }

    private int NormalizeSegmentProbability(int probability) {
        if (probability < 0 || probability > 255) return 128;
        return probability;
    }

    private void WriteSignedLiteral(OfficeVp8BoolEncoder writer, int value, int bits) {
        if (value < 0) {
            writer.WriteLiteral(-value, bits);
            writer.WriteBool(128, true);
        } else {
            writer.WriteLiteral(value, bits);
            writer.WriteBool(128, false);
        }
    }

    private void WriteLoopFilter(OfficeVp8BoolEncoder writer, int baseQIndex, int quality) {
        var level = baseQIndex * 63 / 127;
        if (level < 0) level = 0;
        if (level > 63) level = 63;

        var sharpness = quality >= 85 ? 4 : quality >= 60 ? 2 : 0;
        if (sharpness > 7) sharpness = 7;

        writer.WriteBool(128, false); // filter type (normal)
        writer.WriteLiteral(level, 6);
        writer.WriteLiteral(sharpness, 3);
        writer.WriteBool(128, false); // delta enabled
    }

    private void WriteQuantization(OfficeVp8BoolEncoder writer, int baseQIndex) {
        writer.WriteLiteral(baseQIndex, 7);
        for (var i = 0; i < 5; i++) {
            writer.WriteBool(128, false); // no delta updates
        }
    }

    private void WriteCoefficientProbabilityUpdates(OfficeVp8BoolEncoder writer) {
        var totalCount = CoeffBlockTypes * CoeffBands * CoeffPrevContexts * CoeffEntropyNodes;
        for (var i = 0; i < totalCount; i++) {
            var prob = OfficeVp8Tables.CoeffUpdateProbs[i];
            writer.WriteBool(prob, false);
        }
    }

    private void WriteKeyframeYMode(OfficeVp8BoolEncoder writer, int mode) {
        WriteTree(writer, OfficeVp8Tables.KeyframeYModeTree, OfficeVp8Tables.KeyframeYModeProbs, mode);
    }

    private void WriteKeyframeUvMode(OfficeVp8BoolEncoder writer, int mode) {
        if ((uint)mode >= IntraUvModeCount) mode = ModeDcPred;
        WriteTree(writer, OfficeVp8Tables.UvModeTree, OfficeVp8Tables.KeyframeUvModeProbs, mode);
    }

    private void WriteKeyframeBMode(OfficeVp8BoolEncoder writer, int aboveMode, int leftMode, int mode) {
        if ((uint)aboveMode >= Intra4x4ModeCount) aboveMode = BModeDcPred;
        if ((uint)leftMode >= Intra4x4ModeCount) leftMode = BModeDcPred;
        if ((uint)mode >= Intra4x4ModeCount) mode = BModeDcPred;

        var node = 0;
        while (true) {
            var probIndex = node >> 1;
            var prob = OfficeVp8Tables.KeyframeBModeProbs[((aboveMode * Intra4x4ModeCount) + leftMode) * 9 + probIndex];
            var left = OfficeVp8Tables.BModeTree[node];
            var right = OfficeVp8Tables.BModeTree[node + 1];

            if (ContainsTreeValue(OfficeVp8Tables.BModeTree, left, mode)) {
                writer.WriteBool(prob, false);
                if (left <= 0) return;
                node = left;
            } else {
                writer.WriteBool(prob, true);
                if (right <= 0) return;
                node = right;
            }
        }
    }

    private void WriteTree(OfficeVp8BoolEncoder writer, int[] tree, byte[] probs, int value) {
        var node = 0;
        while (true) {
            var probIndex = node >> 1;
            if ((uint)probIndex >= (uint)probs.Length) return;
            var left = tree[node];
            var right = tree[node + 1];

            if (ContainsTreeValue(tree, left, value)) {
                writer.WriteBool(probs[probIndex], false);
                if (left <= 0) return;
                node = left;
            } else {
                writer.WriteBool(probs[probIndex], true);
                if (right <= 0) return;
                node = right;
            }
        }
    }

    private bool ContainsTreeValue(int[] tree, int nodeValue, int value) {
        if (nodeValue <= 0) {
            return -nodeValue == value;
        }
        var left = tree[nodeValue];
        var right = tree[nodeValue + 1];
        return ContainsTreeValue(tree, left, value) || ContainsTreeValue(tree, right, value);
    }

    private int[] BuildDefaultProbabilities() {
        var totalCount = CoeffBlockTypes * CoeffBands * CoeffPrevContexts * CoeffEntropyNodes;
        var probs = new int[totalCount];
        for (var i = 0; i < totalCount; i++) {
            probs[i] = OfficeVp8Tables.DefaultCoeffProbs[i];
        }
        return probs;
    }

}
