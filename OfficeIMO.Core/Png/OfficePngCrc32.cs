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
        for (int index = offset; index < end; index++) {
            crc = Table[(crc ^ data[index]) & 0xFF] ^ (crc >> 8);
        }
        return crc;
    }

    internal static uint Complete(uint crc) => crc ^ 0xFFFFFFFFU;

    internal static uint Compute(byte[] data, int offset, int count) =>
        Complete(Append(Begin(), data, offset, count));

    private static uint[] CreateTable() {
        var table = new uint[256];
        for (uint index = 0; index < table.Length; index++) {
            uint value = index;
            for (int bit = 0; bit < 8; bit++) {
                value = (value & 1) != 0 ? 0xEDB88320U ^ (value >> 1) : value >> 1;
            }
            table[index] = value;
        }
        return table;
    }
}
