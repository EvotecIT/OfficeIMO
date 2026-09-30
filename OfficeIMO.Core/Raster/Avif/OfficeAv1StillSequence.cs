namespace OfficeIMO.Drawing;

/// <summary>Validated reduced still-picture sequence parameters used by the forthcoming frame/tile decoder.</summary>
internal sealed class OfficeAv1StillSequence {
    internal int MaximumWidth { get; set; }
    internal int MaximumHeight { get; set; }
    internal int WidthBits { get; set; }
    internal int HeightBits { get; set; }
    internal int Level { get; set; }
    internal bool Use128Superblock { get; set; }
    internal bool FilterIntra { get; set; }
    internal bool IntraEdgeFilter { get; set; }
    internal bool SuperResolution { get; set; }
    internal bool Cdef { get; set; }
    internal bool Restoration { get; set; }
    internal bool Monochrome { get; set; }
    internal bool SeparateUvDeltaQ { get; set; }
    internal bool FilmGrain { get; set; }
    internal int ChromaSamplePosition { get; set; }
    internal OfficeAvifColorDescription Color { get; set; } = null!;
    internal int FrameOffset { get; set; }
    internal int FrameLength { get; set; }
}
