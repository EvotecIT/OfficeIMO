using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Tile-local AV1 partition probabilities and completed-block neighbor dimensions.</summary>
/// <remarks>One decoder owns each instance. Record a leaf only after its block syntax is complete.
/// Above widths and left heights replace the specification's two-dimensional MiSizes grid.</remarks>
internal sealed partial class OfficeAv1PartitionContext {
    private readonly OfficeAv1Tile _tile;
    private readonly OfficeAv1TileGeometry _geometry;
    private readonly int _miRows, _miCols;
    private readonly byte[] _aboveWidths, _leftHeights;
    private readonly int[][][] _cdfs = CreateCdfs();
    private readonly int[] _edgeCdf = new int[3];
    private readonly CancellationToken _cancellation;

    internal OfficeAv1PartitionContext(OfficeAv1StillFrame frame, OfficeAv1Tile tile, int superblockPixels,
        CancellationToken cancellation = default) {
        cancellation.ThrowIfCancellationRequested();
        _geometry = new OfficeAv1TileGeometry(frame,tile,superblockPixels);
        _tile = tile; _miRows = frame.MiRows; _miCols = frame.MiCols;
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

    private void ValidateBlock(int row, int col, int width, int height) =>
        _geometry.Validate(new OfficeAv1BlockRegion(row,col,width,height));

    private static int SizeIndex(int pixels) { int index = 0; while (pixels > 8) { pixels >>= 1; index++; } return index; }
    private static int Probability(int[] cdf, int symbol) => cdf[symbol] - cdf[symbol - 1];
}
