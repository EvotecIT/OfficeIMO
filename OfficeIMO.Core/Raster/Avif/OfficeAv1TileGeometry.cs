using System;

namespace OfficeIMO.Drawing;

/// <summary>Validated AV1 4x4-unit tile boundaries and coded block dimensions, including partial frame edges.</summary>
internal readonly struct OfficeAv1TileGeometry {
    internal OfficeAv1TileGeometry(OfficeAv1StillFrame frame, OfficeAv1Tile tile, int superblockPixels) {
        if (frame == null) throw new ArgumentNullException(nameof(frame));
        if (superblockPixels != 64 && superblockPixels != 128) throw new FormatException("Invalid AV1 superblock size.");
        int units = superblockPixels / 4;
        if (frame.MiRows < 2 || frame.MiCols < 2 || frame.MiRows > 16384 || frame.MiCols > 16384 ||
            (frame.MiRows & 1) != 0 || (frame.MiCols & 1) != 0 ||
            tile.MiRowStart < 0 || tile.MiColStart < 0 || tile.MiRowEnd > frame.MiRows || tile.MiColEnd > frame.MiCols ||
            tile.MiRowEnd <= tile.MiRowStart || tile.MiColEnd <= tile.MiColStart ||
            tile.MiRowStart % units != 0 || tile.MiColStart % units != 0 ||
            (tile.MiRowEnd != frame.MiRows && tile.MiRowEnd % units != 0) ||
            (tile.MiColEnd != frame.MiCols && tile.MiColEnd % units != 0))
            throw new FormatException("Invalid AV1 tile geometry.");
        Tile = tile; MiRows = frame.MiRows; MiCols = frame.MiCols; SuperblockPixels = superblockPixels;
    }
    internal OfficeAv1Tile Tile { get; }
    internal int MiRows { get; }
    internal int MiCols { get; }
    internal int SuperblockPixels { get; }

    /// <summary>Checks coded dimensions; only an outer frame edge may truncate the context writes.</summary>
    internal void Validate(OfficeAv1BlockRegion b) {
        if (!Dimension(b.Width) || !Dimension(b.Height) || b.Width > SuperblockPixels || b.Height > SuperblockPixels ||
            Math.Max(b.Width,b.Height) > Math.Min(b.Width,b.Height) * (Math.Max(b.Width,b.Height)==128?2:4) ||
            b.MiRow < Tile.MiRowStart || b.MiRow >= Tile.MiRowEnd || b.MiCol < Tile.MiColStart || b.MiCol >= Tile.MiColEnd ||
            b.MiRow % (b.Height/4) != 0 || b.MiCol % (b.Width/4) != 0 ||
            (b.MiRow+b.Height/4 > Tile.MiRowEnd && Tile.MiRowEnd != MiRows) ||
            (b.MiCol+b.Width/4 > Tile.MiColEnd && Tile.MiColEnd != MiCols))
            throw new FormatException("Invalid AV1 block geometry.");
    }

    /// <summary>Rejects padded geometry if even its minimum possible pixel size exceeds the request limit.</summary>
    internal void EnsurePixelBudget(long maximumPixels) {
        if ((long)(MiRows*4-7)*(MiCols*4-7)>maximumPixels) throw new FormatException("AV1 context exceeds the pixel limit.");
    }
    private static bool Dimension(int pixels) => pixels>=4 && pixels<=128 && (pixels&(pixels-1))==0;
}
