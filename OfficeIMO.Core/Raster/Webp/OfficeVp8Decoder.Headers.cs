// Adapted from CodeGlyphX, commit fc25e2fcf795d9c9a09b88708c47bdfeed5c446d.
// OfficeIMO's copy is licensed under the repository's MIT license by the original author.
// OfficeIMO adaptation removes diagnostic scaffolding and adds bounded cancellation/resource handling.
using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeVp8Decoder {
    internal static bool TryReadHeader(OfficeByteView payload, out OfficeVp8Header header) {
        header = default;
        if (payload.Length < MinimumHeaderBytes) return false;

        var frameTag = ReadU24LE(payload, 0);
        var frameType = frameTag & 1;
        if (frameType != 0) return false; // interframes are not valid for WebP stills

        var version = (frameTag >> 1) & 0x7;
        var showFrame = (frameTag >> 4) & 1;
        var partitionSize = frameTag >> 5;
        if (partitionSize <= 0) return false;

        // Keyframe start code.
        if (payload[3] != 0x9D || payload[4] != 0x01 || payload[5] != 0x2A) return false;

        var widthRaw = ReadU16LE(payload, 6);
        var heightRaw = ReadU16LE(payload, 8);
        var width = widthRaw & 0x3FFF;
        var height = heightRaw & 0x3FFF;
        var horizontalScale = (widthRaw >> 14) & 0x3;
        var verticalScale = (heightRaw >> 14) & 0x3;
        if (width <= 0 || height <= 0) return false;

        header = new OfficeVp8Header(
            width,
            height,
            version,
            showFrame != 0,
            partitionSize,
            horizontalScale,
            verticalScale,
            bitsConsumed: MinimumHeaderBytes * 8);
        return true;
    }

    internal static bool TryGetFirstPartition(OfficeByteView payload, out OfficeByteView firstPartition) {
        firstPartition = default;
        if (!TryReadHeader(payload, out var header)) return false;

        var offset = MinimumHeaderBytes;
        var length = header.PartitionSize;
        if (offset < 0 || length < 0) return false;
        if (offset + length > payload.Length) return false;

        firstPartition = payload.Slice(offset, length);
        return true;
    }

    internal static bool TryGetBoolCodedData(OfficeByteView payload, out OfficeByteView boolData) {
        boolData = default;
        if (!TryGetFirstPartition(payload, out var firstPartition)) return false;
        boolData = firstPartition;
        return true;
    }

    private static bool TryReadFrameHeader(OfficeByteView payload, out OfficeVp8FrameHeader frameHeader, CancellationToken cancellationToken) {
        frameHeader = default;
        if (!TryGetBoolCodedData(payload, out var boolData) || boolData.Length < 2) return false;
        return TryReadFrameHeader(new OfficeVp8BoolDecoder(boolData, cancellationToken), out frameHeader);
    }

    private static bool TryReadFrameHeader(OfficeVp8BoolDecoder decoder, out OfficeVp8FrameHeader frameHeader) {
        frameHeader = default;
        if (!TryReadControlHeader(decoder, out var controlHeader)) return false;
        if (!TryReadSegmentation(decoder, out var segmentation)) return false;
        if (!TryReadLoopFilter(decoder, out var loopFilter)) return false;

        if (!decoder.TryReadLiteral(2, out var partitionBits)) return false;
        if (partitionBits < 0 || partitionBits > 3) return false;
        var dctPartitionCount = 1 << partitionBits;

        if (!TryReadQuantization(decoder, out var quantization)) return false;
        if (!decoder.TryReadBool(128, out var refreshEntropyProbsBit)) return false;
        if (!TryReadCoefficientProbabilities(decoder, out var coefficientProbabilities)) return false;
        if (!decoder.TryReadBool(128, out var noCoefficientSkipBit)) return false;

        var skipProbability = 0;
        if (noCoefficientSkipBit) {
            if (!decoder.TryReadLiteral(8, out skipProbability)) return false;
        }

        frameHeader = new OfficeVp8FrameHeader(
            controlHeader,
            segmentation,
            loopFilter,
            quantization,
            coefficientProbabilities,
            dctPartitionCount,
            refreshEntropyProbsBit,
            noCoefficientSkipBit,
            skipProbability,
            decoder.BytesConsumed);
        return true;
    }

    private static bool TryReadMacroblockHeadersKeyframe(
        OfficeVp8BoolDecoder decoder,
        OfficeVp8Header header,
        OfficeVp8FrameHeader frameHeader,
        out OfficeVp8MacroblockHeader[] macroblocks) {
        macroblocks = Array.Empty<OfficeVp8MacroblockHeader>();

        var macroblockCols = GetMacroblockDimension(header.Width);
        var macroblockRows = GetMacroblockDimension(header.Height);
        if (macroblockCols <= 0 || macroblockRows <= 0) return false;

        var macroblockCountLong = (long)macroblockCols * macroblockRows;
        if (macroblockCountLong <= 0 || macroblockCountLong > int.MaxValue) return false;
        var macroblockCount = (int)macroblockCountLong;

        macroblocks = new OfficeVp8MacroblockHeader[macroblockCount];

        var aboveSubModes = new int[macroblockCols * MacroblockSubBlockCount];
        var currentSubModes = new int[macroblockCols * MacroblockSubBlockCount];

        for (var row = 0; row < macroblockRows; row++) {
            for (var col = 0; col < macroblockCols; col++) {
                var index = (row * macroblockCols) + col;

                var segmentId = 0;
                if (frameHeader.Segmentation.Enabled && frameHeader.Segmentation.UpdateMap) {
                    if (!TryReadSegmentId(decoder, frameHeader.Segmentation.SegmentProbabilities, out segmentId)) {
                        return false;
                    }
                }

                var skipCoefficients = false;
                if (frameHeader.NoCoefficientSkip) {
                    var probability = frameHeader.SkipProbability;
                    if ((uint)probability > 255) return false;
                    if (!decoder.TryReadBool(probability, out skipCoefficients)) return false;
                }

                if (!TryReadKeyframeYMode(decoder, out var yMode)) return false;
                var is4x4 = yMode == YModeBPred;

                var subblockModes = EmptySubblockModes;
                var baseIndex = col * MacroblockSubBlockCount;
                if (is4x4) {
                    subblockModes = new int[MacroblockSubBlockCount];
                    for (var subY = 0; subY < 4; subY++) {
                        for (var subX = 0; subX < 4; subX++) {
                            var blockIndex = (subY * 4) + subX;
                            var aboveMode = BModeDcPred;
                            if (subY == 0) {
                                if (row > 0) {
                                    aboveMode = aboveSubModes[baseIndex + 12 + subX];
                                }
                            } else {
                                aboveMode = subblockModes[((subY - 1) * 4) + subX];
                            }

                            var leftMode = BModeDcPred;
                            if (subX == 0) {
                                if (col > 0) {
                                    leftMode = currentSubModes[baseIndex - MacroblockSubBlockCount + (subY * 4) + 3];
                                }
                            } else {
                                leftMode = subblockModes[(subY * 4) + subX - 1];
                            }

                            if (!TryReadKeyframeBMode(decoder, aboveMode, leftMode, out var mode)) return false;
                            subblockModes[blockIndex] = mode;
                        }
                    }

                    Array.Copy(subblockModes, 0, currentSubModes, baseIndex, MacroblockSubBlockCount);
                } else {
                    for (var i = 0; i < MacroblockSubBlockCount; i++) {
                        currentSubModes[baseIndex + i] = yMode switch {
                            ModeVPred => BModeVEPred,
                            ModeHPred => BModeHEPred,
                            ModeTMPred => BModeTMPred,
                            _ => BModeDcPred
                        };
                    }
                }

                if (!TryReadKeyframeUvMode(decoder, out var uvMode)) return false;

                macroblocks[index] = new OfficeVp8MacroblockHeader(
                    index,
                    col,
                    row,
                    segmentId,
                    skipCoefficients,
                    yMode,
                    uvMode,
                    is4x4,
                    subblockModes);
            }

            SwapRowModes(ref aboveSubModes, ref currentSubModes);
        }

        return true;
    }

    private static bool TryReadKeyframeYMode(OfficeVp8BoolDecoder decoder, out int mode) {
        return TryReadTree(decoder, OfficeVp8Tables.KeyframeYModeTree, OfficeVp8Tables.KeyframeYModeProbs, out mode);
    }

    private static bool TryReadKeyframeUvMode(OfficeVp8BoolDecoder decoder, out int mode) {
        return TryReadTree(decoder, OfficeVp8Tables.UvModeTree, OfficeVp8Tables.KeyframeUvModeProbs, out mode);
    }

    private static bool TryReadKeyframeBMode(OfficeVp8BoolDecoder decoder, int aboveMode, int leftMode, out int mode) {
        mode = 0;
        if ((uint)aboveMode >= Intra4x4ModeCount) aboveMode = BModeDcPred;
        if ((uint)leftMode >= Intra4x4ModeCount) leftMode = BModeDcPred;

        var node = 0;
        while (true) {
            var probIndex = node >> 1;
            if ((uint)probIndex >= 9) return false;
            var prob = OfficeVp8Tables.KeyframeBModeProbs[((aboveMode * Intra4x4ModeCount) + leftMode) * 9 + probIndex];
            if (!decoder.TryReadBool(prob, out var bit)) return false;
            var nextIndex = node + (bit ? 1 : 0);
            if ((uint)nextIndex >= (uint)OfficeVp8Tables.BModeTree.Length) return false;
            var next = OfficeVp8Tables.BModeTree[nextIndex];
            if (next <= 0) {
                mode = -next;
                return true;
            }
            node = next;
        }
    }

    private static bool TryReadTree(
        OfficeVp8BoolDecoder decoder,
        int[] tree,
        byte[] probs,
        out int value) {
        value = 0;
        var node = 0;
        while (true) {
            var probIndex = node >> 1;
            if ((uint)probIndex >= (uint)probs.Length) return false;
            if (!decoder.TryReadBool(probs[probIndex], out var bit)) return false;
            var nextIndex = node + (bit ? 1 : 0);
            if ((uint)nextIndex >= (uint)tree.Length) return false;
            var next = tree[nextIndex];
            if (next <= 0) {
                value = -next;
                return true;
            }
            node = next;
        }
    }

    internal static bool TryReadPartitionLayout(OfficeByteView payload, out OfficeVp8PartitionLayout layout, CancellationToken cancellationToken) {
        layout = default;
        if (!TryReadHeader(payload, out var header)) return false;
        if (!TryReadFrameHeader(payload, out var frameHeader, cancellationToken)) return false;

        var firstPartitionOffset = MinimumHeaderBytes;
        var firstPartitionSize = header.PartitionSize;
        if (firstPartitionSize <= 0) return false;

        var firstPartitionEndLong = (long)firstPartitionOffset + firstPartitionSize;
        if (firstPartitionEndLong > payload.Length) return false;
        if (firstPartitionEndLong > int.MaxValue) return false;
        var firstPartitionEnd = (int)firstPartitionEndLong;

        var dctPartitionCount = frameHeader.DctPartitionCount;
        if (dctPartitionCount <= 0) return false;

        var sizeTableBytesLong = 3L * (dctPartitionCount - 1);
        if (sizeTableBytesLong > int.MaxValue) return false;
        var sizeTableBytes = (int)sizeTableBytesLong;
        var sizeTableOffset = firstPartitionEnd;
        var dctDataOffsetLong = (long)sizeTableOffset + sizeTableBytes;
        if (dctDataOffsetLong > payload.Length) return false;
        if (dctDataOffsetLong > int.MaxValue) return false;
        var dctDataOffset = (int)dctDataOffsetLong;

        var dctSizes = new int[dctPartitionCount];
        long sum = 0;
        for (var i = 0; i < dctPartitionCount - 1; i++) {
            var entryOffset = sizeTableOffset + (i * 3);
            var size = ReadU24LE(payload, entryOffset);
            if (size < 0) return false;
            dctSizes[i] = size;
            sum += size;
            if (sum > payload.Length) return false;
        }

        var remainingLong = payload.Length - dctDataOffset - sum;
        if (remainingLong < 0) return false;
        if (remainingLong > int.MaxValue) return false;
        dctSizes[dctPartitionCount - 1] = (int)remainingLong;

        layout = new OfficeVp8PartitionLayout(
            firstPartitionOffset,
            firstPartitionSize,
            sizeTableOffset,
            sizeTableBytes,
            dctDataOffset,
            dctSizes,
            payload.Length - dctDataOffset);
        return true;
    }

    private static bool TryReadControlHeader(OfficeVp8BoolDecoder decoder, out OfficeVp8ControlHeader controlHeader) {
        controlHeader = default;
        if (!decoder.TryReadBool(probability: 128, out var colorSpaceBit)) return false;
        if (!decoder.TryReadBool(probability: 128, out var clampTypeBit)) return false;

        controlHeader = new OfficeVp8ControlHeader(
            colorSpaceBit ? 1 : 0,
            clampTypeBit ? 1 : 0,
            decoder.BytesConsumed);
        return true;
    }

    private static bool TryReadSegmentation(OfficeVp8BoolDecoder decoder, out OfficeVp8Segmentation segmentation) {
        segmentation = default;
        if (!decoder.TryReadBool(128, out var segmentationEnabledBit)) return false;

        var enabled = segmentationEnabledBit;
        var updateMap = false;
        var updateData = false;
        var absoluteDeltas = false;
        var quantizerDeltas = new int[SegmentCount];
        var filterDeltas = new int[SegmentCount];
        var segmentProbabilities = new int[SegmentProbCount];

        for (var i = 0; i < segmentProbabilities.Length; i++) {
            segmentProbabilities[i] = -1;
        }

        if (enabled) {
            if (!decoder.TryReadBool(128, out var updateMapBit)) return false;
            if (!decoder.TryReadBool(128, out var updateDataBit)) return false;
            updateMap = updateMapBit;
            updateData = updateDataBit;

            if (updateData) {
                if (!decoder.TryReadBool(128, out var absoluteDeltasBit)) return false;
                absoluteDeltas = absoluteDeltasBit;

                for (var i = 0; i < SegmentCount; i++) {
                    if (!decoder.TryReadBool(128, out var updateQuantizerBit)) return false;
                    if (updateQuantizerBit) {
                        if (!decoder.TryReadSignedLiteral(7, out quantizerDeltas[i])) return false;
                    }
                }

                for (var i = 0; i < SegmentCount; i++) {
                    if (!decoder.TryReadBool(128, out var updateFilterBit)) return false;
                    if (updateFilterBit) {
                        if (!decoder.TryReadSignedLiteral(6, out filterDeltas[i])) return false;
                    }
                }
            }

            if (updateMap) {
                for (var i = 0; i < SegmentProbCount; i++) {
                    if (!decoder.TryReadBool(128, out var updateProbabilityBit)) return false;
                    if (updateProbabilityBit) {
                        if (!decoder.TryReadLiteral(8, out segmentProbabilities[i])) return false;
                    }
                }
            }
        }

        segmentation = new OfficeVp8Segmentation(
            enabled,
            updateMap,
            updateData,
            absoluteDeltas,
            quantizerDeltas,
            filterDeltas,
            segmentProbabilities);
        return true;
    }

    private static bool TryReadQuantization(OfficeVp8BoolDecoder decoder, out OfficeVp8Quantization quantization) {
        quantization = default;
        if (!decoder.TryReadLiteral(7, out var baseQIndex)) return false;
        if (baseQIndex < 0 || baseQIndex > 127) return false;

        var deltas = new int[QuantDeltaCount];
        var deltasUpdated = new bool[QuantDeltaCount];

        for (var i = 0; i < QuantDeltaCount; i++) {
            if (!decoder.TryReadBool(128, out var updateDeltaBit)) return false;
            if (updateDeltaBit) {
                if (!decoder.TryReadSignedLiteral(4, out deltas[i])) return false;
                deltasUpdated[i] = true;
            }
        }

        quantization = new OfficeVp8Quantization(baseQIndex, deltas, deltasUpdated);
        return true;
    }

    private static bool TryReadCoefficientProbabilities(OfficeVp8BoolDecoder decoder, out OfficeVp8CoefficientProbabilities probabilities) {
        probabilities = default;
        var totalCount = CoeffBlockTypes * CoeffBands * CoeffPrevContexts * CoeffEntropyNodes;
        var probs = new int[totalCount];
        var updated = new bool[totalCount];

        for (var i = 0; i < probs.Length; i++) {
            probs[i] = OfficeVp8Tables.DefaultCoeffProbs[i];
        }

        var updatedCount = 0;
        for (var blockType = 0; blockType < CoeffBlockTypes; blockType++) {
            for (var band = 0; band < CoeffBands; band++) {
                for (var prev = 0; prev < CoeffPrevContexts; prev++) {
                    for (var node = 0; node < CoeffEntropyNodes; node++) {
                        var index = GetCoeffIndex(blockType, band, prev, node);
                        var updateProbability = OfficeVp8Tables.CoeffUpdateProbs[index];
                        if (!decoder.TryReadBool(updateProbability, out var updateBit)) return false;
                        if (!updateBit) continue;

                        if (!decoder.TryReadLiteral(8, out var value)) return false;
                        probs[index] = value;
                        updated[index] = true;
                        updatedCount++;
                    }
                }
            }
        }

        probabilities = new OfficeVp8CoefficientProbabilities(probs, updated, updatedCount, decoder.BytesConsumed);
        return true;
    }

    private static bool TryReadCoefficientToken(
        OfficeVp8BoolDecoder decoder,
        OfficeVp8CoefficientProbabilities probabilities,
        int blockType,
        int band,
        int prevContext,
        out int tokenCode, bool skipEndOfBlock = false) {
        tokenCode = 0;
        if ((uint)blockType >= CoeffBlockTypes) return false;
        if ((uint)band >= CoeffBands) return false;
        if ((uint)prevContext >= CoeffPrevContexts) return false;

        // A zero token cannot be followed by end-of-block.
        var node = skipEndOfBlock ? 2 : 0;
        while (node >= 0) {
            var probabilityIndex = node >> 1;
            if ((uint)probabilityIndex >= CoeffEntropyNodes) return false;

            var coeffIndex = GetCoeffIndex(blockType, band, prevContext, node: probabilityIndex);
            var probability = probabilities.Probabilities[coeffIndex];
            if ((uint)probability > 255) return false;
            if (!decoder.TryReadBool(probability, out var bit)) return false;

            var nextIndex = node + (bit ? 1 : 0);
            if ((uint)nextIndex >= CoeffTokenTree.Length) return false;
            node = CoeffTokenTree[nextIndex];
        }

        var token = -node - 1;
        if ((uint)token >= CoeffTokenCount) return false;
        tokenCode = token;
        return true;
    }

    private static bool TryReadSegmentId(OfficeVp8BoolDecoder decoder, int[] probabilities, out int segmentId) {
        segmentId = 0;
        if (probabilities.Length < SegmentProbCount) return false;

        int root = NormalizeSegmentProbability(probabilities[0]);
        if (!decoder.TryReadBool(root, out bool upperPair)) return false;
        int branch = NormalizeSegmentProbability(probabilities[upperPair ? 2 : 1]);
        if (!decoder.TryReadBool(branch, out bool upperMember)) return false;
        segmentId = (upperPair ? 2 : 0) + (upperMember ? 1 : 0);
        return true;
    }

    private static int NormalizeSegmentProbability(int probability) {
        return probability < 0 || probability > 255 ? 255 : probability;
    }

    private static int GetCoeffIndex(int blockType, int band, int prev, int node) {
        return (((blockType * CoeffBands) + band) * CoeffPrevContexts + prev) * CoeffEntropyNodes + node;
    }

    private static int ReadU16LE(OfficeByteView data, int offset) {
        if (offset < 0 || offset + 2 > data.Length) return 0;
        return data[offset] | (data[offset + 1] << 8);
    }

    private static int ReadU24LE(OfficeByteView data, int offset) {
        if (offset < 0 || offset + 3 > data.Length) return 0;
        return data[offset]
            | (data[offset + 1] << 8)
            | (data[offset + 2] << 16);
    }
}
