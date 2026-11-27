namespace OfficeIMO.Archives.Compression;

/// <summary>CRC-64 with the ECMA-182 polynomial in reflected form, as the xz format uses it.</summary>
internal static class Crc64 {
    private const ulong Polynomial = 0xC96C5795D7870F42UL;
    private static readonly ulong[] Table = CreateTable();

    internal static ulong Begin() => ulong.MaxValue;

    internal static ulong Append(ulong crc, byte[] data, int offset, int count) {
        for (int index = offset, end = offset + count; index < end; index++)
            crc = Table[(byte)(crc ^ data[index])] ^ (crc >> 8);
        return crc;
    }

    internal static ulong Complete(ulong crc) => ~crc;

    internal static ulong Compute(byte[] data, int offset, int count) => Complete(Append(Begin(), data, offset, count));

    private static ulong[] CreateTable() {
        var table = new ulong[256];
        for (uint value = 0; value < 256; value++) {
            ulong crc = value;
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? (crc >> 1) ^ Polynomial : crc >> 1;
            table[value] = crc;
        }
        return table;
    }
}
