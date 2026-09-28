// Adapted from CodeGlyphX, commit fc25e2fcf795d9c9a09b88708c47bdfeed5c446d.
// Copyright CodeGlyphX contributors. Apache-2.0; see THIRD-PARTY-NOTICES.md.
// OfficeIMO adaptation removes diagnostic scaffolding and adds bounded cancellation/resource handling.
using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal readonly struct OfficeVp8Header {
    public OfficeVp8Header(
        int width,
        int height,
        int version,
        bool showFrame,
        int partitionSize,
        int horizontalScale,
        int verticalScale,
        int bitsConsumed) {
        Width = width;
        Height = height;
        Version = version;
        ShowFrame = showFrame;
        PartitionSize = partitionSize;
        HorizontalScale = horizontalScale;
        VerticalScale = verticalScale;
        BitsConsumed = bitsConsumed;
    }

    public int Width { get; }
    public int Height { get; }
    public int Version { get; }
    public bool ShowFrame { get; }
    public int PartitionSize { get; }
    public int HorizontalScale { get; }
    public int VerticalScale { get; }
    public int BitsConsumed { get; }
}

internal readonly struct OfficeVp8DecodeStats {
    public OfficeVp8DecodeStats(int macroblockCount, int macroblocksWithCoefficients, int skippedMacroblocks) {
        MacroblockCount = macroblockCount;
        MacroblocksWithCoefficients = macroblocksWithCoefficients;
        SkippedMacroblocks = skippedMacroblocks;
    }

    public int MacroblockCount { get; }
    public int MacroblocksWithCoefficients { get; }
    public int SkippedMacroblocks { get; }
}

internal readonly struct OfficeVp8ControlHeader {
    public OfficeVp8ControlHeader(int colorSpace, int clampType, int bytesConsumed) {
        ColorSpace = colorSpace;
        ClampType = clampType;
        BytesConsumed = bytesConsumed;
    }

    public int ColorSpace { get; }
    public int ClampType { get; }
    public int BytesConsumed { get; }
}

internal readonly struct OfficeVp8Segmentation {
    public OfficeVp8Segmentation(
        bool enabled,
        bool updateMap,
        bool updateData,
        bool absoluteDeltas,
        int[] quantizerDeltas,
        int[] filterDeltas,
        int[] segmentProbabilities) {
        Enabled = enabled;
        UpdateMap = updateMap;
        UpdateData = updateData;
        AbsoluteDeltas = absoluteDeltas;
        QuantizerDeltas = quantizerDeltas;
        FilterDeltas = filterDeltas;
        SegmentProbabilities = segmentProbabilities;
    }

    public bool Enabled { get; }
    public bool UpdateMap { get; }
    public bool UpdateData { get; }
    public bool AbsoluteDeltas { get; }
    public int[] QuantizerDeltas { get; }
    public int[] FilterDeltas { get; }
    public int[] SegmentProbabilities { get; }
}

internal readonly struct OfficeVp8LoopFilter {
    public OfficeVp8LoopFilter(
        int filterType,
        int level,
        int sharpness,
        bool deltaEnabled,
        bool deltaUpdate,
        int[] refDeltas,
        bool[] refDeltasUpdated,
        int[] modeDeltas,
        bool[] modeDeltasUpdated) {
        FilterType = filterType;
        Level = level;
        Sharpness = sharpness;
        DeltaEnabled = deltaEnabled;
        DeltaUpdate = deltaUpdate;
        RefDeltas = refDeltas;
        RefDeltasUpdated = refDeltasUpdated;
        ModeDeltas = modeDeltas;
        ModeDeltasUpdated = modeDeltasUpdated;
    }

    public int FilterType { get; }
    public int Level { get; }
    public int Sharpness { get; }
    public bool DeltaEnabled { get; }
    public bool DeltaUpdate { get; }
    public int[] RefDeltas { get; }
    public bool[] RefDeltasUpdated { get; }
    public int[] ModeDeltas { get; }
    public bool[] ModeDeltasUpdated { get; }
}

internal readonly struct OfficeVp8Quantization {
    public OfficeVp8Quantization(int baseQIndex, int[] deltas, bool[] deltasUpdated) {
        BaseQIndex = baseQIndex;
        Deltas = deltas;
        DeltasUpdated = deltasUpdated;
    }

    public int BaseQIndex { get; }
    public int[] Deltas { get; }
    public bool[] DeltasUpdated { get; }
}

internal readonly struct OfficeVp8CoefficientProbabilities {
    public OfficeVp8CoefficientProbabilities(int[] probabilities, bool[] updated, int updatedCount, int bytesConsumed) {
        Probabilities = probabilities;
        Updated = updated;
        UpdatedCount = updatedCount;
        BytesConsumed = bytesConsumed;
    }

    public int[] Probabilities { get; }
    public bool[] Updated { get; }
    public int UpdatedCount { get; }
    public int BytesConsumed { get; }
}

internal readonly struct OfficeVp8FrameHeader {
    public OfficeVp8FrameHeader(
        OfficeVp8ControlHeader controlHeader,
        OfficeVp8Segmentation segmentation,
        OfficeVp8LoopFilter loopFilter,
        OfficeVp8Quantization quantization,
        OfficeVp8CoefficientProbabilities coefficientProbabilities,
        int dctPartitionCount,
        bool refreshEntropyProbs,
        bool noCoefficientSkip,
        int skipProbability,
        int bytesConsumed) {
        ControlHeader = controlHeader;
        Segmentation = segmentation;
        LoopFilter = loopFilter;
        Quantization = quantization;
        CoefficientProbabilities = coefficientProbabilities;
        DctPartitionCount = dctPartitionCount;
        RefreshEntropyProbs = refreshEntropyProbs;
        NoCoefficientSkip = noCoefficientSkip;
        SkipProbability = skipProbability;
        BytesConsumed = bytesConsumed;
    }

    public OfficeVp8ControlHeader ControlHeader { get; }
    public OfficeVp8Segmentation Segmentation { get; }
    public OfficeVp8LoopFilter LoopFilter { get; }
    public OfficeVp8Quantization Quantization { get; }
    public OfficeVp8CoefficientProbabilities CoefficientProbabilities { get; }
    public int DctPartitionCount { get; }
    public bool RefreshEntropyProbs { get; }
    public bool NoCoefficientSkip { get; }
    public int SkipProbability { get; }
    public int BytesConsumed { get; }
}

internal readonly struct OfficeVp8PartitionLayout {
    public OfficeVp8PartitionLayout(
        int firstPartitionOffset,
        int firstPartitionSize,
        int sizeTableOffset,
        int sizeTableBytes,
        int dctDataOffset,
        int[] dctPartitionSizes,
        int dctBytesAvailable) {
        FirstPartitionOffset = firstPartitionOffset;
        FirstPartitionSize = firstPartitionSize;
        SizeTableOffset = sizeTableOffset;
        SizeTableBytes = sizeTableBytes;
        DctDataOffset = dctDataOffset;
        DctPartitionSizes = dctPartitionSizes;
        DctBytesAvailable = dctBytesAvailable;
    }

    public int FirstPartitionOffset { get; }
    public int FirstPartitionSize { get; }
    public int SizeTableOffset { get; }
    public int SizeTableBytes { get; }
    public int DctDataOffset { get; }
    public int[] DctPartitionSizes { get; }
    public int DctBytesAvailable { get; }
}

internal readonly struct OfficeVp8TokenInfo {
    public OfficeVp8TokenInfo(
        int tokenCode,
        int band,
        int prevContextBefore,
        int prevContextAfter,
        bool hasMore,
        bool isNonZero) {
        TokenCode = tokenCode;
        Band = band;
        PrevContextBefore = prevContextBefore;
        PrevContextAfter = prevContextAfter;
        HasMore = hasMore;
        IsNonZero = isNonZero;
    }

    public int TokenCode { get; }
    public int Band { get; }
    public int PrevContextBefore { get; }
    public int PrevContextAfter { get; }
    public bool HasMore { get; }
    public bool IsNonZero { get; }
}

internal readonly struct OfficeVp8MacroblockHeader {
    public OfficeVp8MacroblockHeader(
        int index,
        int x,
        int y,
        int segmentId,
        bool skipCoefficients,
        int yMode,
        int uvMode,
        bool is4x4,
        int[] subblockModes) {
        Index = index;
        X = x;
        Y = y;
        SegmentId = segmentId;
        SkipCoefficients = skipCoefficients;
        YMode = yMode;
        UvMode = uvMode;
        Is4x4 = is4x4;
        SubblockModes = subblockModes ?? Array.Empty<int>();
    }

    public int Index { get; }
    public int X { get; }
    public int Y { get; }
    public int SegmentId { get; }
    public bool SkipCoefficients { get; }
    public int YMode { get; }
    public int UvMode { get; }
    public bool Is4x4 { get; }
    public int[] SubblockModes { get; }
}
