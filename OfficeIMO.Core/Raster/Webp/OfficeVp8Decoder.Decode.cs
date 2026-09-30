// Adapted from CodeGlyphX, commit fc25e2fcf795d9c9a09b88708c47bdfeed5c446d.
// OfficeIMO's copy is licensed under the repository's MIT license by the original author.
// OfficeIMO adaptation removes diagnostic scaffolding and adds bounded cancellation/resource handling.
using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeVp8Decoder {
    private const int MinimumHeaderBytes = FrameTagBytes + StartCodeBytes + DimensionBytes;
    private const int SegmentCount = 4;
    private const int SegmentProbCount = 3;
    private const int LoopFilterDeltaCount = 4;
    private const int QuantDeltaCount = 5;
    private const int CoeffBlockTypes = 4;
    private const int CoeffBands = 8;
    private const int CoeffPrevContexts = 3;
    private const int CoeffEntropyNodes = 11;
    private const int CoefficientsPerBlock = 16;
    private const int CoeffTokenCount = 12;
    private const int Intra4x4ModeCount = 10;
    private const int YModeBPred = 4;
    private const int ModeVPred = 1;
    private const int ModeHPred = 2;
    private const int ModeTMPred = 3;
    private const int BModeDcPred = 0;
    private const int BModeTMPred = 1;
    private const int BModeVEPred = 2;
    private const int BModeHEPred = 3;
    private const int MacroblockSubBlockCount = 16;
    private const int MacroblockSize = 16;
    private const int BlockSize = 4;
    private const int MacroblockChromaWidth = MacroblockSize / 2;
    private const int MacroblockChromaHeight = MacroblockSize / 2;
    private const int MacroblockChromaBlocks = 4;
    private static readonly int[] CoeffBandTable = new[]
    {
        0, 1, 2, 3,
        6, 4, 5, 6,
        6, 6, 6, 6,
        6, 6, 6, 7,
    };
    private static readonly int[] ZigZagToNaturalOrder = new[]
    {
        0,  1,  4,  8,
        5,  2,  3,  6,
        9, 12, 13, 10,
        7, 11, 14, 15,
    };
    private static readonly int[] CoeffTokenExtraBits = new[]
    {
        0, 0, 0, 0, 0, 0,
        1, 2, 3, 4, 5, 11,
    };
    private static readonly int[] CoeffTokenBaseMagnitude = new[]
    {
        0, 0, 1, 2, 3, 4,
        5, 7, 11, 19, 35, 67,
    };
    private static readonly int[] CoeffTokenTree = new[] {
        -1, 2, -2, 4, -3, 6, 8, 12, -4, 10, -5, -6,
        14, 16, -7, -8, 18, 20, -9, -10, -11, -12
    };
    private static readonly int[] EmptySubblockModes = Array.Empty<int>();
    private static readonly int[] ZeroCoefficients = new int[CoefficientsPerBlock];
    private const int FrameTagBytes = 3;
    private const int StartCodeBytes = 3;
    private const int DimensionBytes = 4;

    private readonly struct DequantFactors {
        public DequantFactors(int y1Dc, int y1Ac, int y2Dc, int y2Ac, int uvDc, int uvAc) {
            Y1Dc = y1Dc;
            Y1Ac = y1Ac;
            Y2Dc = y2Dc;
            Y2Ac = y2Ac;
            UvDc = uvDc;
            UvAc = uvAc;
        }

        public int Y1Dc { get; }
        public int Y1Ac { get; }
        public int Y2Dc { get; }
        public int Y2Ac { get; }
        public int UvDc { get; }
        public int UvAc { get; }
    }

    internal static bool TryDecode(OfficeByteView payload, CancellationToken cancellationToken, long retainedManagedBytes, out byte[] rgba, out int width, out int height) {
        rgba = Array.Empty<byte>();
        width = 0;
        height = 0;

        cancellationToken.ThrowIfCancellationRequested();
        if (!TryReadHeader(payload, out var header) || header.Version > 3 || !header.ShowFrame ||
            !OfficeRasterGuards.TryEnsurePixelCount(header.Width, header.Height, out int pixelCount)) return false;
        long paddedWidth = GetMacroblockDimension(header.Width) * 16L;
        long paddedHeight = GetMacroblockDimension(header.Height) * 16L;
        long macroblocks = paddedWidth * paddedHeight / 256L;
        // Simultaneously retained planes, RGBA output, macroblock headers/modes,
        // arithmetic partition copies, row contexts and bounded scratch arrays.
        long requiredBytes = checked(retainedManagedBytes + payload.Length * 3L +
            paddedWidth * paddedHeight * 3L / 2L + pixelCount * 4L + macroblocks * 192L + paddedWidth * 16L + 65536L);
        if (requiredBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
        if (!TryDecodeKeyframe(payload, header, cancellationToken, out rgba, out _)) return false;

        width = header.Width;
        height = header.Height;
        return true;
    }

    private static bool TryDecodeKeyframe(
        OfficeByteView payload,
        OfficeVp8Header header,
        CancellationToken cancellationToken,
        out byte[] rgba,
        out OfficeVp8DecodeStats stats) {
        rgba = Array.Empty<byte>();
        stats = default;

        if (!TryGetBoolCodedData(payload, out var boolData)) return false;
        var headerDecoder = new OfficeVp8BoolDecoder(boolData, cancellationToken);
        if (!TryReadFrameHeader(headerDecoder, out var frameHeader)) return false;

        if (!TryReadMacroblockHeadersKeyframe(headerDecoder, header, frameHeader, out var macroblocks)) return false;

        if (!TryReadPartitionLayout(payload, out var layout, cancellationToken)) return false;
        var partitionCount = layout.DctPartitionSizes.Length;
        if (partitionCount <= 0) return false;
        if (macroblocks.Length == 0) return false;

        var tokenDecoders = new OfficeVp8BoolDecoder[partitionCount];
        var offset = layout.DctDataOffset;
        for (var i = 0; i < partitionCount; i++) {
            var size = layout.DctPartitionSizes[i];
            if (size < 0) return false;
            if (offset < 0 || offset + size > payload.Length) return false;
            tokenDecoders[i] = new OfficeVp8BoolDecoder(payload.Slice(offset, size), cancellationToken);
            offset += size;
        }

        var width = GetMacroblockDimension(header.Width) * MacroblockSize;
        var height = GetMacroblockDimension(header.Height) * MacroblockSize;
        var yPlane = new byte[checked(width * height)];
        var chromaWidth = (width + 1) >> 1;
        var chromaHeight = (height + 1) >> 1;
        var uPlane = new byte[checked(chromaWidth * chromaHeight)];
        var vPlane = new byte[checked(chromaWidth * chromaHeight)];

        var macroblockCols = GetMacroblockDimension(width);
        var macroblockRows = GetMacroblockDimension(height);
        if (macroblockCols <= 0 || macroblockRows <= 0) return false;
        if (macroblocks.Length != macroblockCols * macroblockRows) return false;

        var dequant = BuildDequantFactors(frameHeader.Quantization, frameHeader.Segmentation);

        var y2NzAbove = new byte[macroblockCols];
        var y2NzCurrent = new byte[macroblockCols];
        var yNzAbove = new byte[macroblockCols * MacroblockSubBlockCount];
        var yNzCurrent = new byte[macroblockCols * MacroblockSubBlockCount];
        var uNzAbove = new byte[macroblockCols * MacroblockChromaBlocks];
        var uNzCurrent = new byte[macroblockCols * MacroblockChromaBlocks];
        var vNzAbove = new byte[macroblockCols * MacroblockChromaBlocks];
        var vNzCurrent = new byte[macroblockCols * MacroblockChromaBlocks];

        var macroblockHasCoefficients = new bool[macroblocks.Length];
        var predictionScratch = new OfficeVp8DecodeScratch();

        for (var row = 0; row < macroblockRows; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            var partitionIndex = GetTokenPartitionForRow(row, partitionCount);
            if ((uint)partitionIndex >= (uint)tokenDecoders.Length) return false;
            var tokenDecoder = tokenDecoders[partitionIndex];
            byte y2Left = 0;

            for (var col = 0; col < macroblockCols; col++) {
                var macroblockIndex = (row * macroblockCols) + col;
                var macroblock = macroblocks[macroblockIndex];
                var segmentId = macroblock.SegmentId;
                if ((uint)segmentId >= (uint)dequant.Length) segmentId = 0;

                var hasCoefficients = false;
                if (!TryDecodeMacroblock(
                    tokenDecoder,
                    frameHeader.CoefficientProbabilities,
                    macroblock,
                    dequant[segmentId],
                    col,
                    row,
                    width,
                    height,
                    yPlane,
                    uPlane,
                    vPlane,
                    chromaWidth,
                    chromaHeight,
                    y2NzAbove,
                    y2NzCurrent,
                    ref y2Left,
                    yNzAbove,
                    yNzCurrent,
                    uNzAbove,
                    uNzCurrent,
                    vNzAbove,
                    vNzCurrent,
                    predictionScratch,
                    out hasCoefficients)) {
                    return false;
                }

                macroblockHasCoefficients[macroblockIndex] = hasCoefficients;
            }

            SwapRowContexts(ref y2NzAbove, ref y2NzCurrent);
            SwapRowContexts(ref yNzAbove, ref yNzCurrent);
            SwapRowContexts(ref uNzAbove, ref uNzCurrent);
            SwapRowContexts(ref vNzAbove, ref vNzCurrent);
        }

        var macroblocksWithCoefficients = 0;
        var skippedMacroblocks = 0;
        for (var i = 0; i < macroblocks.Length; i++) {
            if (macroblockHasCoefficients[i]) {
                macroblocksWithCoefficients++;
            }
            if (macroblocks[i].SkipCoefficients) {
                skippedMacroblocks++;
            }
        }

        ApplyLoopFilter(
            frameHeader.LoopFilter,
            frameHeader.Segmentation,
            macroblocks,
            macroblockHasCoefficients,
            width,
            height,
            yPlane,
            uPlane,
            vPlane,
            chromaWidth,
            chromaHeight,
            isKeyframe: true, cancellationToken);

        rgba = ConvertPlanesToRgba(header.Width, header.Height, width, yPlane, uPlane, vPlane, chromaWidth, chromaHeight, cancellationToken);
        stats = new OfficeVp8DecodeStats(macroblocks.Length, macroblocksWithCoefficients, skippedMacroblocks);
        return true;
    }
}
