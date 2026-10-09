// Adapted from CodeGlyphX commit 3c9103da07623e15804450ec0bca284bfcbf4a2c.
// Copyright (c) Przemyslaw Klys. Licensed under Apache-2.0; see Licenses/CodeGlyphX-LICENSE.txt.
// OfficeIMO adaptation adds bounded output, cancellation, and shared prediction/reconstruction.
using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// Managed VP8 lossy encoder (minimal intra-only encoder).
/// </summary>
internal sealed partial class OfficeVp8Encoder {
    private const int BlockSize = 4;
    private const int MacroblockSize = 16;
    private const int MacroblockSubBlockCount = 16;
    private const int MacroblockChromaBlocks = 4;
    private const int Intra4x4ModeCount = 10;
    private const int IntraUvModeCount = 4;
    private const int SegmentCount = 4;
    private const int SegmentProbCount = 3;
    private const int YModeBPred = 4;
    private const int BModeDcPred = 0;
    private const int ModeDcPred = 0;
    private const int BlockTypeY2 = OfficeVp8Tables.LumaSecondOrder;
    private const int BlockTypeY = OfficeVp8Tables.LumaWithDc;
    private const int BlockTypeU = OfficeVp8Tables.Chroma;
    private const int BlockTypeV = OfficeVp8Tables.Chroma;
    private const int CoeffBlockTypes = 4;
    private const int CoeffBands = 8;
    private const int CoeffPrevContexts = 3;
    private const int CoeffEntropyNodes = 11;
    private const int CoefficientsPerBlock = 16;
    private const int MaxCoefficientMagnitude = 2047;

    private static readonly int[] CoeffBandTable =
    {
        0, 1, 2, 3,
        6, 4, 5, 6,
        6, 6, 6, 6,
        6, 6, 6, 7,
    };

    private static readonly int[] ZigZagToNaturalOrder =
    {
        0, 1, 4, 8,
        5, 2, 3, 6,
        9, 12, 13, 10,
        7, 11, 14, 15,
    };

    private static readonly int[] CoeffTokenExtraBits =
    {
        0, 0, 0, 0, 0, 0,
        1, 2, 3, 4, 5, 11,
    };

    private static readonly int[] CoeffTokenBaseMagnitude =
    {
        0, 0, 1, 2, 3, 4,
        5, 7, 11, 19, 35, 67,
    };

    private static readonly int[] CoeffTokenTree = {
        -1, 2, -2, 4, -3, 6, 8, 12, -4, 10, -5, -6,
        14, 16, -7, -8, 18, 20, -9, -10, -11, -12
    };

    private readonly CancellationToken _cancellationToken;
    private readonly Action<OfficeRasterEncodingCheckpoint>? _checkpointObserver;
    private readonly OfficeVp8DecodeScratch _scratch = new OfficeVp8DecodeScratch();
    private readonly int _maximumPayloadBytes;
    private readonly OfficeVp8EncodingMemory _memory;

    private OfficeVp8Encoder(CancellationToken cancellationToken, int maximumPayloadBytes, OfficeVp8EncodingMemory memory,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        _cancellationToken = cancellationToken;
        _maximumPayloadBytes = maximumPayloadBytes;
        _memory = memory;
        _checkpointObserver = checkpointObserver;
    }

    internal static void EncodeRgba32(byte[] rgba, int width, int height, int quality,
        int maximumPayloadBytes, OfficeVp8EncodingMemory memory, CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver,
        out byte[] payload, out byte[] alphaPayload) {
        var encoder = new OfficeVp8Encoder(cancellationToken, maximumPayloadBytes, memory, checkpointObserver);
        encoder.Checkpoint();
        if (!encoder.TryEncodeVp8Payload(rgba, width, height, checked(width * 4), quality, out payload, out string reason)) {
            throw new NotSupportedException(reason);
        }
        alphaPayload = encoder.ComputeAlphaUsed(rgba, width, height, width * 4)
            ? encoder.BuildAlphPayload(rgba, width, height, width * 4) : Array.Empty<byte>();
        encoder.Checkpoint();
    }

    private void Checkpoint() {
        _checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.WebpCompressionBlock);
        _cancellationToken.ThrowIfCancellationRequested();
    }

    private bool TryEncodeVp8Payload(
        byte[] rgba,
        int width,
        int height,
        int stride,
        int quality,
        out byte[] payload,
        out string reason) {
        payload = Array.Empty<byte>();
        reason = string.Empty;

        int visibleWidth = width, visibleHeight = height;
        var baseQIndex = QualityToBaseQIndex(quality);

        ConvertRgbaToYuv420(rgba, width, height, stride, out var yPlane, out var uPlane, out var vPlane);

        int paddedWidth = GetMacroblockDimension(width) * MacroblockSize;
        int paddedHeight = GetMacroblockDimension(height) * MacroblockSize;
        yPlane = PadPlane(yPlane, width, height, paddedWidth, paddedHeight);
        uPlane = PadPlane(uPlane, (width + 1) >> 1, (height + 1) >> 1, paddedWidth / 2, paddedHeight / 2);
        vPlane = PadPlane(vPlane, (width + 1) >> 1, (height + 1) >> 1, paddedWidth / 2, paddedHeight / 2);
        width = paddedWidth; height = paddedHeight;
        var reconY = new byte[yPlane.Length];
        var reconU = new byte[uPlane.Length];
        var reconV = new byte[vPlane.Length];

        var macroblockCols = GetMacroblockDimension(width);
        var macroblockRows = GetMacroblockDimension(height);
        if (macroblockCols <= 0 || macroblockRows <= 0) {
            reason = "Invalid macroblock geometry.";
            return false;
        }

        var segmentIds = BuildSegmentMap(yPlane, width, height, macroblockCols, macroblockRows, baseQIndex, out var segmentationEnabled, out var quantizerDeltas, out var segmentProbabilities);
        var dequantSegments = BuildDequantFactors(baseQIndex, quantizerDeltas, segmentationEnabled);

        var headerWriter = new OfficeVp8BoolEncoder(Math.Min(4096, _maximumPayloadBytes), Math.Min(0x7FFFF, _maximumPayloadBytes), _cancellationToken, _memory);
        WriteControlHeader(headerWriter);
        WriteSegmentation(headerWriter, segmentationEnabled, quantizerDeltas, segmentProbabilities);
        WriteLoopFilter(headerWriter, baseQIndex, quality);
        headerWriter.WriteLiteral(0, 2); // one DCT partition
        WriteQuantization(headerWriter, baseQIndex);
        headerWriter.WriteBool(128, false); // refresh entropy probs
        WriteCoefficientProbabilityUpdates(headerWriter);
        var enableSkip = true;
        var skipProbability = ComputeSkipProbability(baseQIndex, segmentationEnabled);
        headerWriter.WriteBool(128, enableSkip);
        if (enableSkip) {
            headerWriter.WriteLiteral(skipProbability, 8);
        }

        var tokenWriter = new OfficeVp8BoolEncoder(Math.Min(4096, _maximumPayloadBytes), _maximumPayloadBytes, _cancellationToken, _memory);

        var aboveSubModes = new int[macroblockCols * MacroblockSubBlockCount];
        var currentSubModes = new int[macroblockCols * MacroblockSubBlockCount];
        var yNzAbove = new byte[macroblockCols * MacroblockSubBlockCount];
        var yNzCurrent = new byte[macroblockCols * MacroblockSubBlockCount];
        var y2NzAbove = new byte[macroblockCols];
        var y2NzCurrent = new byte[macroblockCols];
        var uNzAbove = new byte[macroblockCols * MacroblockChromaBlocks];
        var uNzCurrent = new byte[macroblockCols * MacroblockChromaBlocks];
        var vNzAbove = new byte[macroblockCols * MacroblockChromaBlocks];
        var vNzCurrent = new byte[macroblockCols * MacroblockChromaBlocks];

        var probabilities = BuildDefaultProbabilities();

        for (var row = 0; row < macroblockRows; row++) {
            Checkpoint();
            byte y2Left = 0;
            for (var col = 0; col < macroblockCols; col++) {
                var macroblockIndex = (row * macroblockCols) + col;
                var segmentId = segmentationEnabled && segmentIds.Length > macroblockIndex ? segmentIds[macroblockIndex] : 0;
                EncodeMacroblock(
                    headerWriter,
                    tokenWriter,
                    probabilities,
                    yPlane,
                    uPlane,
                    vPlane,
                    reconY,
                    reconU,
                    reconV,
                    width,
                    height,
                    col,
                    row,
                    dequantSegments,
                    segmentId,
                    y2NzAbove,
                    y2NzCurrent,
                    ref y2Left,
                    yNzAbove,
                    yNzCurrent,
                    uNzAbove,
                    uNzCurrent,
                    vNzAbove,
                    vNzCurrent,
                    aboveSubModes,
                    currentSubModes,
                    segmentationEnabled,
                    segmentProbabilities,
                    enableSkip,
                    skipProbability);
            }

            SwapRowModes(ref aboveSubModes, ref currentSubModes);
            SwapRowContexts(ref y2NzAbove, ref y2NzCurrent);
            SwapRowContexts(ref yNzAbove, ref yNzCurrent);
            SwapRowContexts(ref uNzAbove, ref uNzCurrent);
            SwapRowContexts(ref vNzAbove, ref vNzCurrent);
        }

        var headerBytes = headerWriter.Finish();
        if (headerBytes.Length < 2) headerBytes = PadToLength(headerBytes, 2);
        var tokenBytes = tokenWriter.Finish();
        if (tokenBytes.Length < 2) tokenBytes = PadToLength(tokenBytes, 2);

        var keyframeHeader = BuildKeyframeHeader(visibleWidth, visibleHeight);
        _memory.Reserve(keyframeHeader.Length + headerBytes.Length + 24L);
        var firstPartition = new byte[keyframeHeader.Length + headerBytes.Length];
        Buffer.BlockCopy(keyframeHeader, 0, firstPartition, 0, keyframeHeader.Length);
        Buffer.BlockCopy(headerBytes, 0, firstPartition, keyframeHeader.Length, headerBytes.Length);

        Checkpoint();
        var firstPartitionSize = headerBytes.Length;
        if (firstPartitionSize > 0x7FFFF) {
            reason = "VP8 first partition is too large.";
            return false;
        }

        var frameTag = BuildFrameTag(firstPartitionSize, version: 0, showFrame: true, keyframe: true);

        int totalBytes = checked(frameTag.Length + firstPartition.Length + tokenBytes.Length);
        if (totalBytes > _maximumPayloadBytes) throw new ArgumentException("VP8 output exceeds encoded-size limits.");
        _memory.Reserve(totalBytes + 24L);
        payload = new byte[totalBytes];
        Buffer.BlockCopy(frameTag, 0, payload, 0, frameTag.Length);
        Buffer.BlockCopy(firstPartition, 0, payload, frameTag.Length, firstPartition.Length);
        Buffer.BlockCopy(tokenBytes, 0, payload, frameTag.Length + firstPartition.Length, tokenBytes.Length);

        return true;
    }

}
