namespace OfficeIMO.Archives.Tests;

internal sealed partial class ArchiveFixtures {
    internal static byte[] LzxCab(bool literal, byte[]? trailing = null, bool corruptPadding = false) {
        var bits = new List<int>();
        void Bits(int value, int count) {
            for (int index = count - 1; index >= 0; index--) bits.Add((value >> index) & 1);
        }
        void Tree(int count, bool literals) {
            for (int symbol = 0; symbol < 20; symbol++) Bits(symbol == 0 || symbol == 16 ? 1 : 0, 4);
            for (int symbol = 0; symbol < count; symbol++) Bits(literals && (symbol == 'A' || symbol == 'B') ? 1 : 0, 1);
        }
        Bits(0, 1); Bits(literal ? 1 : 3, 3); Bits(0, 16); Bits(4, 8);
        if (literal) { Tree(256, true); Tree(256, false); Tree(249, false); Bits(0, 4); }
        while (bits.Count % 16 != 0) bits.Add(0);
        using var chunk = new MemoryStream();
        using (var writer = new BinaryWriter(chunk, Encoding.UTF8, leaveOpen: true)) {
            for (int offset = 0; offset < bits.Count; offset += 16) {
                ushort word = 0;
                for (int index = 0; index < 16; index++) word = (ushort)((word << 1) | bits[offset + index]);
                writer.Write(word);
            }
            if (!literal) { writer.Write(1u); writer.Write(1u); writer.Write(1u); writer.Write(Encoding.ASCII.GetBytes("ABCD")); }
            if (trailing != null) writer.Write(trailing);
        }
        byte[] payload = chunk.ToArray();
        if (corruptPadding) payload[payload.Length - 2] |= 1;
        byte[] cabinet = StoredCab("Template.admx", payload);
        int dataOffset = (int)BitConverter.ToUInt32(cabinet, 36);
        Buffer.BlockCopy(BitConverter.GetBytes((ushort)((16 << 8) | 3)), 0, cabinet, 42, 2);
        Buffer.BlockCopy(BitConverter.GetBytes(4u), 0, cabinet, 44, 4);
        Buffer.BlockCopy(BitConverter.GetBytes((ushort)4), 0, cabinet, dataOffset + 6, 2);
        return cabinet;
    }
}
