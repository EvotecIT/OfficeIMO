using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1PaletteReader {
    private ushort[] ReadColors(OfficeAv1SymbolReader symbols, OfficeAv1BlockRegion b, int plane, int count) {
        int above = (b.MiRow & 15) != 0 ? AboveSize(plane, b) : 0, left = LeftSize(plane, b);
        int n = 0;
        for (int i = 0; i < above; i++) _cache[n++] = _aboveColors[(b.MiCol - _geometry.Tile.MiColStart) * 16 + plane * 8 + i];
        for (int i = 0; i < left; i++) _cache[n++] = _leftColors[(b.MiRow - _geometry.Tile.MiRowStart) * 16 + plane * 8 + i];
        Array.Sort(_cache, 0, n);
        int unique = 0;
        for (int i = 0; i < n; i++) if (unique == 0 || _cache[i] != _cache[unique - 1]) _cache[unique++] = _cache[i];
        ushort[] colors = new ushort[count]; int index = 0;
        for (int i = 0; i < unique && index < count; i++) if (symbols.ReadBool()) colors[index++] = _cache[i];
        if (index < count) colors[index++] = (ushort)symbols.ReadLiteral(8);
        int bits = index < count ? 5 + symbols.ReadLiteral(2) : 0;
        while (index < count) {
            int delta = symbols.ReadLiteral(bits) + (plane == 0 ? 1 : 0);
            colors[index] = (ushort)Math.Min(255, colors[index - 1] + delta);
            int range = 256 - colors[index] - (plane == 0 ? 1 : 0);
            bits = Math.Min(bits, range <= 1 ? 0 : Log2(range - 1) + 1);
            index++;
        }
        Array.Sort(colors);
        return colors;
    }

    private static ushort[] ReadVColors(OfficeAv1SymbolReader symbols, int count) {
        ushort[] colors = new ushort[count];
        if (!symbols.ReadBool()) {
            for (int i = 0; i < count; i++) colors[i] = (ushort)symbols.ReadLiteral(8);
        } else {
            int bits = 4 + symbols.ReadLiteral(2);
            colors[0] = (ushort)symbols.ReadLiteral(8);
            for (int i = 1; i < count; i++) {
                int delta = symbols.ReadLiteral(bits);
                if (delta != 0 && symbols.ReadBool()) delta = -delta;
                colors[i] = (ushort)((colors[i - 1] + delta + 256) & 255);
            }
        }
        return colors;
    }
}
