using System;

namespace OfficeIMO.Drawing;

internal static class OfficePngCrc32 {
    private static readonly uint[] Table = CreateTable();

    internal static uint Begin() => 0xFFFFFFFFU;

    internal static uint Append(uint crc, byte[] data, int offset, int count) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (offset < 0 || count < 0 || offset > data.Length - count) {
            throw new ArgumentOutOfRangeException(nameof(offset));
        }
        int end = offset + count;
        int index = offset;
        // Eight successive reflected CRC steps can be evaluated together.
        // Explicit byte assembly retains the same ordering on every target.
        for (; index <= end - 8; index += 8) {
            uint first = crc ^ ((uint)data[index] | ((uint)data[index + 1] << 8) |
                ((uint)data[index + 2] << 16) | ((uint)data[index + 3] << 24));
            uint second = (uint)data[index + 4] | ((uint)data[index + 5] << 8) |
                ((uint)data[index + 6] << 16) | ((uint)data[index + 7] << 24);
            crc = Table[1792 + (first & 0xFF)] ^ Table[1536 + ((first >> 8) & 0xFF)] ^
                Table[1280 + ((first >> 16) & 0xFF)] ^ Table[1024 + (first >> 24)] ^
                Table[768 + (second & 0xFF)] ^ Table[512 + ((second >> 8) & 0xFF)] ^
                Table[256 + ((second >> 16) & 0xFF)] ^ Table[second >> 24];
        }
        for (; index < end; index++) {
            crc = Table[(crc ^ data[index]) & 0xFF] ^ (crc >> 8);
        }
        return crc;
    }

    internal static uint Complete(uint crc) => crc ^ 0xFFFFFFFFU;

    internal static uint Compute(byte[] data, int offset, int count) =>
        Complete(Append(Begin(), data, offset, count));

    private static uint[] CreateTable() {
        var table = new uint[8 * 256];
        for (uint index = 0; index < 256; index++) {
            uint value = index;
            for (int bit = 0; bit < 8; bit++) {
                value = (value & 1) != 0 ? 0xEDB88320U ^ (value >> 1) : value >> 1;
            }
            table[index] = value;
        }
        for (int generation = 1; generation < 8; generation++) {
            for (int index = 0; index < 256; index++) {
                uint value = table[(generation - 1) * 256 + index];
                table[generation * 256 + index] = table[value & 0xFF] ^ (value >> 8);
            }
        }
        return table;
    }
}
