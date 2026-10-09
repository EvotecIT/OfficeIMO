namespace OfficeIMO.DjVu;

// DjVu v3 section 8.3.10: BGR palette and BZZ-compressed, big-endian placement indices.
internal sealed class DjVuPalette {
    private readonly byte[] _colors, _indices;
    private DjVuPalette(byte[] colors, byte[] indices) { _colors = colors; _indices = indices; }
    internal void Paint(int placement, byte[] output, int offset) {
        int index = DjVuBinary.U16(_indices, placement * 2) * 3;
        output[offset] = _colors[index + 2]; output[offset + 1] = _colors[index + 1]; output[offset + 2] = _colors[index]; output[offset + 3] = 255;
    }

    internal static DjVuPalette Decode(DjVuChunk chunk, int placements, DjVuReadBudget budget) {
        if (chunk.Length < 3) throw new InvalidDataException("Truncated DjVu foreground palette.");
        byte[] data = chunk.Source;
        int offset = chunk.Offset, end = offset + chunk.Length, version = data[offset++];
        if ((version & 127) != 0) throw new NotSupportedException("Unsupported DjVu palette version.");
        int count = DjVuBinary.U16(data, offset); offset += 2;
        if (count == 0 || count * 3 > end - offset) throw new InvalidDataException("Invalid DjVu palette size.");
        budget.WorkingBytes(count * 3L + placements * 2L);
        byte[] colors = new byte[count * 3];
        Buffer.BlockCopy(data, offset, colors, 0, colors.Length); offset += colors.Length;
        if ((version & 128) == 0 || end - offset < 3) throw new InvalidDataException("DjVu foreground palette has no placement correspondence.");
        if (DjVuBinary.U24(data, offset) != placements) throw new InvalidDataException("DjVu palette placement count differs from the JB2 mask.");
        offset += 3;
        budget.RetainBytes(colors.LongLength);
        byte[] indices = BzzDecoder.Decode(data, offset, end - offset, budget);
        budget.WorkingBytes(indices.LongLength);
        if (indices.LongLength != placements * 2L) throw new InvalidDataException("Invalid DjVu palette correspondence length.");
        for (int i = 0; i < placements; i++) {
            if ((i & 4095) == 0) budget.Cancellation.ThrowIfCancellationRequested();
            if (DjVuBinary.U16(indices, i * 2) >= count) throw new InvalidDataException("DjVu palette index is outside the color table.");
        }
        budget.RetainBytes(indices.LongLength);
        return new DjVuPalette(colors, indices);
    }
}
