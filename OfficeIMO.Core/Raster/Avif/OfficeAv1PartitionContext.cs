using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Tile-local AV1 partition probabilities and completed-block neighbor dimensions.</summary>
/// <remarks>One decoder owns each instance. Record a leaf only after its block syntax is complete.
/// Above widths and left heights replace the specification's two-dimensional MiSizes grid.</remarks>
internal sealed partial class OfficeAv1PartitionContext {
    private readonly OfficeAv1Tile _tile;
    private readonly int _miRows, _miCols, _superblockPixels;
    private readonly byte[] _aboveWidths, _leftHeights;
    private readonly int[][][] _cdfs = CreateCdfs();
    private readonly int[] _edgeCdf = new int[3];
    private readonly CancellationToken _cancellation;

    internal OfficeAv1PartitionContext(OfficeAv1StillFrame frame, OfficeAv1Tile tile, int superblockPixels,
        CancellationToken cancellation = default) {
        cancellation.ThrowIfCancellationRequested();
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
            throw new FormatException("Invalid AV1 partition tile geometry.");
        _tile = tile; _miRows = frame.MiRows; _miCols = frame.MiCols; _superblockPixels = superblockPixels;
        _cancellation = cancellation;
        _aboveWidths = new byte[tile.MiColEnd - tile.MiColStart];
        _leftHeights = new byte[tile.MiRowEnd - tile.MiRowStart];
    }

    /// <summary>Reads one square partition node. Leaf mode/coefficient syntax is decoded separately before recording it.</summary>
    internal OfficeAv1PartitionLayout Read(OfficeAv1SymbolReader symbols, int row, int col, int pixels) {
        _cancellation.ThrowIfCancellationRequested();
        if (symbols == null) throw new ArgumentNullException(nameof(symbols));
        ValidateBlock(row, col, pixels, pixels);
        OfficeAv1Partition kind;
        int units = pixels / 4, half = units / 2;
        if (pixels == 4) kind = OfficeAv1Partition.None;
        else {
            bool hasRows = row + half < _miRows, hasCols = col + half < _miCols;
            if (!hasRows && !hasCols) kind = OfficeAv1Partition.Split;
            else {
                int context = (row > _tile.MiRowStart && _aboveWidths[col - _tile.MiColStart] < units ? 1 : 0) +
                    (col > _tile.MiColStart && _leftHeights[row - _tile.MiRowStart] < units ? 2 : 0);
                int[] cdf = _cdfs[SizeIndex(pixels)][context];
                if (hasRows && hasCols) kind = (OfficeAv1Partition)symbols.ReadSymbol(cdf);
                else {
                    // AV1 8.3.2: collapse forbidden partitions into split. This temporary CDF never replaces cdf.
                    int splitProbability = hasCols
                        ? Probability(cdf, 2) + Probability(cdf, 3) + Probability(cdf, 4) + Probability(cdf, 6) + Probability(cdf, 7)
                        : Probability(cdf, 1) + Probability(cdf, 3) + Probability(cdf, 4) + Probability(cdf, 5) + Probability(cdf, 6);
                    if (pixels != 128) splitProbability += Probability(cdf, hasCols ? 9 : 8);
                    _edgeCdf[0] = 32768 - splitProbability; _edgeCdf[1] = 32768; _edgeCdf[2] = 0;
                    kind = symbols.ReadSymbol(_edgeCdf) != 0 ? OfficeAv1Partition.Split
                        : hasCols ? OfficeAv1Partition.Horizontal : OfficeAv1Partition.Vertical;
                }
            }
        }
        return new OfficeAv1PartitionLayout(kind, row, col, pixels);
    }

    /// <summary>Retains coded dimensions, clipping context writes at the frame's partial edge.</summary>
    internal void RecordBlock(OfficeAv1BlockRegion block) {
        _cancellation.ThrowIfCancellationRequested();
        ValidateBlock(block.MiRow, block.MiCol, block.Width, block.Height);
        int width = block.Width / 4, height = block.Height / 4;
        int endCol = Math.Min(_tile.MiColEnd, block.MiCol + width);
        int endRow = Math.Min(_tile.MiRowEnd, block.MiRow + height);
        for (int c = block.MiCol; c < endCol; c++) _aboveWidths[c - _tile.MiColStart] = (byte)width;
        for (int r = block.MiRow; r < endRow; r++) _leftHeights[r - _tile.MiRowStart] = (byte)height;
    }

    private void ValidateBlock(int row, int col, int width, int height) {
        if (!IsBlockDimension(width) || !IsBlockDimension(height) || width > _superblockPixels || height > _superblockPixels ||
            Math.Max(width, height) > Math.Min(width, height) * (Math.Max(width, height) == 128 ? 2 : 4) ||
            row < _tile.MiRowStart || row >= _tile.MiRowEnd || col < _tile.MiColStart || col >= _tile.MiColEnd ||
            row % (height / 4) != 0 || col % (width / 4) != 0 ||
            (row + height / 4 > _tile.MiRowEnd && _tile.MiRowEnd != _miRows) ||
            (col + width / 4 > _tile.MiColEnd && _tile.MiColEnd != _miCols))
            throw new FormatException("Invalid AV1 partition block geometry.");
    }

    private static bool IsBlockDimension(int pixels) => pixels >= 4 && pixels <= 128 && (pixels & (pixels - 1)) == 0;
    private static int SizeIndex(int pixels) { int index = 0; while (pixels > 8) { pixels >>= 1; index++; } return index; }
    private static int Probability(int[] cdf, int symbol) => cdf[symbol] - cdf[symbol - 1];
}
