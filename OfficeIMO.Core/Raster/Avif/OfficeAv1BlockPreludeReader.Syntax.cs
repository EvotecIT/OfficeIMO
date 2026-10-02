using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1BlockPreludeReader {
    private int ReadSegment(OfficeAv1SymbolReader symbols, OfficeAv1BlockRegion block, bool skip) {
        if (!_segmentation) return 0;
        int r = block.MiRow, c = block.MiCol;
        bool above = r > _tile.MiRowStart, left = c > _tile.MiColStart;
        int ul = above && left ? _segments![SegmentIndex(r - 1, c - 1)] : -1;
        int u = above ? _segments![SegmentIndex(r - 1, c)] : -1;
        int l = left ? _segments![SegmentIndex(r, c - 1)] : -1;
        int prediction = u == -1 ? l == -1 ? 0 : l : l == -1 ? u : ul == u ? u : l;
        int segment = prediction;
        if (!skip) {
            int context = ul < 0 ? 0 : ul == u && ul == l ? 2 : ul == u || ul == l || u == l ? 1 : 0;
            int diff = symbols.ReadSymbol(_segmentCdfs[context]);
            int maximum = _lastSegment + 1;
            if (prediction == 0) segment = diff;
            else if (prediction >= maximum - 1) segment = maximum - diff - 1;
            else if (2 * prediction < maximum) segment = diff <= 2 * prediction
                ? (diff & 1) != 0 ? prediction + ((diff + 1) >> 1) : prediction - (diff >> 1) : diff;
            else segment = diff <= 2 * (maximum - prediction - 1)
                ? (diff & 1) != 0 ? prediction + ((diff + 1) >> 1) : prediction - (diff >> 1) : maximum - (diff + 1);
        }
        if (segment < 0 || segment > _lastSegment) throw new FormatException("Invalid AV1 segment identifier.");
        return segment;
    }

    private int SegmentIndex(int row, int col) => (row - _tile.MiRowStart) * _columns + col - _tile.MiColStart;
    private int CdefSlot(int row, int col) => ((row - _superblockRow) / 16) * 2 + (col - _superblockCol) / 16;

    private void ReadCdef(OfficeAv1SymbolReader symbols, OfficeAv1BlockRegion block, bool skip) {
        if (!_cdef || skip) return;
        int slot = CdefSlot(block.MiRow, block.MiCol);
        if (_cdefIndices[slot] >= 0) return;
        int index = symbols.ReadLiteral(_cdefBits);
        int row = block.MiRow & ~15, col = block.MiCol & ~15;
        for (int r = row; r < row + block.Height / 4; r += 16)
            for (int c = col; c < col + block.Width / 4; c += 16) _cdefIndices[CdefSlot(r, c)] = index;
    }

    private static int ReadDelta(OfficeAv1SymbolReader symbols, int[] cdf) {
        int absolute = symbols.ReadSymbol(cdf);
        if (absolute == 3) {
            int bits = symbols.ReadLiteral(3) + 1;
            absolute = symbols.ReadLiteral(bits) + (1 << bits) + 1;
        }
        return absolute != 0 && symbols.ReadBool() ? -absolute : absolute;
    }
}
