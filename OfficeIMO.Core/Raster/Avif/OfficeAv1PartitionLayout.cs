using System;

namespace OfficeIMO.Drawing;

/// <summary>AV1 partition symbols, in bitstream order.</summary>
internal enum OfficeAv1Partition { None, Horizontal, Vertical, Split, HorizontalA, HorizontalB, VerticalA, VerticalB, Horizontal4, Vertical4 }

/// <summary>One coded leaf rectangle or recursive square, positioned in 4x4 units with pixel dimensions.</summary>
internal readonly struct OfficeAv1BlockRegion {
    internal OfficeAv1BlockRegion(int miRow, int miCol, int width, int height) {
        MiRow = miRow; MiCol = miCol; Width = width; Height = height;
    }
    internal int MiRow { get; }
    internal int MiCol { get; }
    internal int Width { get; }
    internal int Height { get; }
}

/// <summary>Allocation-free partition child geometry in AV1 decode order.</summary>
/// <remarks>For Split, children are recursive partition nodes. Otherwise they are leaves.
/// The tile decoder skips children whose origin is outside the frame before reading block syntax.</remarks>
internal readonly struct OfficeAv1PartitionLayout {
    internal OfficeAv1PartitionLayout(OfficeAv1Partition kind, int row, int col, int pixels) {
        if (pixels < 4 || pixels > 128 || (pixels & (pixels - 1)) != 0 || row < 0 || col < 0 ||
            (int)kind < 0 || (int)kind > 9 || (pixels == 4 && kind != OfficeAv1Partition.None) ||
            (pixels == 8 && (int)kind > 3) || (pixels == 128 && (int)kind > 7))
            throw new FormatException("Invalid AV1 partition layout.");
        Kind = kind; _row = row; _col = col; _pixels = pixels;
    }
    private readonly int _row, _col, _pixels;
    internal OfficeAv1Partition Kind { get; }
    internal int Count => Kind == OfficeAv1Partition.None ? 1
        : Kind == OfficeAv1Partition.Horizontal || Kind == OfficeAv1Partition.Vertical ? 2
        : Kind == OfficeAv1Partition.Split || Kind == OfficeAv1Partition.Horizontal4 || Kind == OfficeAv1Partition.Vertical4 ? 4 : 3;

    internal OfficeAv1BlockRegion Child(int index) {
        if (index < 0 || index >= Count) throw new ArgumentOutOfRangeException(nameof(index));
        int half = _pixels / 2, halfMi = half / 4, quarter = _pixels / 4, quarterMi = quarter / 4;
        switch (Kind) {
            case OfficeAv1Partition.None: return Block(0, 0, _pixels, _pixels);
            case OfficeAv1Partition.Horizontal: return Block(index * halfMi, 0, _pixels, half);
            case OfficeAv1Partition.Vertical: return Block(0, index * halfMi, half, _pixels);
            case OfficeAv1Partition.Split: return Block(index / 2 * halfMi, index % 2 * halfMi, half, half);
            case OfficeAv1Partition.HorizontalA:
                return index < 2 ? Block(0, index * halfMi, half, half) : Block(halfMi, 0, _pixels, half);
            case OfficeAv1Partition.HorizontalB:
                return index == 0 ? Block(0, 0, _pixels, half) : Block(halfMi, (index - 1) * halfMi, half, half);
            case OfficeAv1Partition.VerticalA:
                return index < 2 ? Block(index * halfMi, 0, half, half) : Block(0, halfMi, half, _pixels);
            case OfficeAv1Partition.VerticalB:
                return index == 0 ? Block(0, 0, half, _pixels) : Block((index - 1) * halfMi, halfMi, half, half);
            case OfficeAv1Partition.Horizontal4: return Block(index * quarterMi, 0, _pixels, quarter);
            default: return Block(0, index * quarterMi, quarter, _pixels);
        }
    }
    private OfficeAv1BlockRegion Block(int row, int col, int width, int height) => new OfficeAv1BlockRegion(_row + row, _col + col, width, height);
}
