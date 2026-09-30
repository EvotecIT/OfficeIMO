using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1PaletteReader {
    private byte[] ReadMap(OfficeAv1SymbolReader symbols, int count, int width, int height, int onWidth, int onHeight, int[][] cdfs) {
        byte[] map = new byte[width * height];
        map[0] = (byte)symbols.ReadNonSymmetric(count);
        for (int diagonal = 1; diagonal < onWidth + onHeight - 1; diagonal++) {
            _cancellation.ThrowIfCancellationRequested();
            for (int col = Math.Min(diagonal, onWidth - 1); col >= Math.Max(0, diagonal - onHeight + 1); col--) {
                int row = diagonal - col;
                int context = ColorContext(map, width, row, col, count);
                map[row * width + col] = (byte)_order[symbols.ReadSymbol(cdfs[(count - 2) * 5 + context])];
            }
        }
        for (int row = 0; row < onHeight; row++)
            for (int col = onWidth; col < width; col++) map[row * width + col] = map[row * width + onWidth - 1];
        for (int row = onHeight; row < height; row++) Array.Copy(map, (onHeight - 1) * width, map, row * width, width);
        return map;
    }

    private int ColorContext(byte[] map, int width, int row, int col, int count) {
        Array.Clear(_scores, 0, 8);
        for (int i = 0; i < 8; i++) _order[i] = i;
        if (col > 0) _scores[map[row * width + col - 1]] += 2;
        if (row > 0 && col > 0) _scores[map[(row - 1) * width + col - 1]]++;
        if (row > 0) _scores[map[(row - 1) * width + col]] += 2;
        for (int i = 0; i < 3; i++) {
            int max = i;
            for (int j = i + 1; j < count; j++) if (_scores[j] > _scores[max]) max = j;
            int score = _scores[max], color = _order[max];
            for (int j = max; j > i; j--) { _scores[j] = _scores[j - 1]; _order[j] = _order[j - 1]; }
            _scores[i] = score; _order[i] = color;
        }
        int hash = _scores[0] + 2 * _scores[1] + 2 * _scores[2];
        switch (hash) {
            case 2: return 0;
            case 5: return 4;
            case 6: return 3;
            case 7: return 2;
            case 8: return 1;
            default: throw new FormatException("Invalid AV1 palette color context.");
        }
    }
}
