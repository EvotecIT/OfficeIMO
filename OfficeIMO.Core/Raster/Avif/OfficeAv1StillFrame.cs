using System;

namespace OfficeIMO.Drawing;

/// <summary>Shown still-key-frame syntax retained for tile reconstruction, without allocating pixel planes.</summary>
internal sealed class OfficeAv1StillFrame {
    internal int BitDepth { get; set; } = 8;
    internal int Width { get; set; }
    internal int UpscaledWidth { get; set; }
    internal int Height { get; set; }
    internal int RenderWidth { get; set; }
    internal int RenderHeight { get; set; }
    internal int SuperResolutionDenominator { get; set; } = 8;
    internal int MiCols { get; set; }
    internal int MiRows { get; set; }
    internal bool DisableCdfUpdate { get; set; }
    internal bool AllowScreenContentTools { get; set; }
    internal bool AllowIntraBlockCopy { get; set; }
    internal int HeaderBytes { get; set; }
    internal int TileGroupHeaderBytes { get; set; }
    internal int TileColsLog2 { get; set; }
    internal int TileRowsLog2 { get; set; }
    internal int[] MiColStarts { get; set; } = Array.Empty<int>();
    internal int[] MiRowStarts { get; set; } = Array.Empty<int>();
    internal int ContextUpdateTileId { get; set; }
    internal int TileSizeBytes { get; set; }
    internal OfficeAv1Tile[] Tiles { get; set; } = Array.Empty<OfficeAv1Tile>();
    internal int BaseQIndex { get; set; }
    internal int DeltaQYDc { get; set; }
    internal int DeltaQUDc { get; set; }
    internal int DeltaQUAc { get; set; }
    internal int DeltaQVDc { get; set; }
    internal int DeltaQVAc { get; set; }
    internal bool UsingQMatrix { get; set; }
    internal int[] QMatrixLevels { get; } = new int[3];
    internal bool SegmentationEnabled { get; set; }
    internal bool[,] SegmentFeatures { get; } = new bool[8, 8];
    internal int[,] SegmentData { get; } = new int[8, 8];
    internal int LastActiveSegmentId { get; set; }
    internal bool SegmentIdPreSkip { get; set; }
    internal bool[] LosslessSegments { get; } = new bool[8];
    internal bool CodedLossless { get; set; }
    internal bool AllLossless { get; set; }
    internal bool DeltaQPresent { get; set; }
    internal int DeltaQResolution { get; set; }
    internal bool DeltaLoopFilterPresent { get; set; }
    internal int DeltaLoopFilterResolution { get; set; }
    internal bool DeltaLoopFilterMulti { get; set; }
    internal int[] LoopFilterLevels { get; } = new int[4];
    internal int LoopFilterSharpness { get; set; }
    internal bool LoopFilterDeltaEnabled { get; set; } = true;
    internal int[] LoopFilterReferenceDeltas { get; } = new[] { 1, 0, 0, 0, -1, 0, -1, -1 };
    internal int[] LoopFilterModeDeltas { get; } = new int[2];
    internal int CdefDamping { get; set; } = 3;
    internal int CdefBits { get; set; }
    internal int[,] CdefStrengths { get; } = new int[8, 4];
    // AV1 restoration enum order: none, switchable, Wiener, self-guided projection.
    internal int[] RestorationTypes { get; } = new int[3];
    internal int[] RestorationUnitSizes { get; } = new int[3];
    // ONLY_4X4 = 0, TX_MODE_LARGEST = 1, TX_MODE_SELECT = 2.
    internal int TransformMode { get; set; }
    internal bool ReducedTransformSet { get; set; }
}

/// <summary>One bounded entropy payload and its 4x4-unit row/column extent.</summary>
internal readonly struct OfficeAv1Tile {
    internal OfficeAv1Tile(int offset, int length, int rowStart, int rowEnd, int colStart, int colEnd) {
        Offset = offset; Length = length; MiRowStart = rowStart; MiRowEnd = rowEnd; MiColStart = colStart; MiColEnd = colEnd;
    }
    internal int Offset { get; }
    internal int Length { get; }
    internal int MiRowStart { get; }
    internal int MiRowEnd { get; }
    internal int MiColStart { get; }
    internal int MiColEnd { get; }
}
