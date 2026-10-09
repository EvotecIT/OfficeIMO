namespace OfficeIMO.Visio;

internal static class VisioBinaryChunks {
    internal static IEnumerable<Chunk> Read(VisioBinaryContainer.Node node, VisioBinaryContainer container, int maxDepth) {
        byte[] data = node.Data;
        int offset = 0;
        while (offset < data.Length) {
            container.CountRecord();
            // Native streams contain zero separators between records.
            while (offset < data.Length && data[offset] == 0) {
                if ((offset & 4095) == 0) container.CheckCancellation();
                offset++;
            }
            if (offset == data.Length) yield break;
            VisioBinaryData.Require(data, offset, 19);
            uint type = VisioBinaryData.U32(data, offset);
            uint id = VisioBinaryData.U32(data, offset + 4);
            uint list = VisioBinaryData.U32(data, offset + 8);
            int length = VisioBinaryData.Size(data, offset + 12);
            int level = VisioBinaryData.U16(data, offset + 16);
            byte flags = data[offset + 18];
            if (level > maxDepth) throw new InvalidDataException("Binary Visio record depth exceeded.");
            int trailer = Trailer(type, list, level, flags);
            int start = offset + 19;
            VisioBinaryData.Require(data, start, length);
            VisioBinaryData.Require(data, start + length, trailer);
            yield return new Chunk(type, id, level, data, start, length);
            offset = start + length + trailer;
        }
    }

    private static int Trailer(uint type, uint list, int level, byte flags) {
        int trailer = list != 0 || type is 0x71 or 0x70 or 0x6b or 0x6a or 0x69 or 0x66 or 0x65 or 0x2c ? 8 : 0;
        if (list != 0 || (level == 2 && flags == 0x55) || (level == 2 && flags == 0x54 && type == 0xaa)
            || (level == 3 && flags is not (0x50 or 0x54))) trailer += 4;
        if (type is 0x64 or 0x65 or 0x66 or 0x69 or 0x6a or 0x6b or 0x6f or 0x71 or 0x92 or 0xa9 or 0xb4 or 0xb6 or 0xb9 or 0xc7) {
            if (trailer is not (12 or 4)) trailer += 4;
        }
        return type is 0x1f or 0xc9 or 0x2d or 0xd1 ? 0 : trailer;
    }

    internal readonly struct Chunk {
        internal Chunk(uint type, uint id, int level, byte[] data, int offset, int length) {
            Type = type; Id = id; Level = level; Data = data; Offset = offset; Length = length;
        }
        internal uint Type { get; }
        internal uint Id { get; }
        internal int Level { get; }
        internal byte[] Data { get; }
        internal int Offset { get; }
        internal int Length { get; }
        internal byte Byte(int offset) { Check(offset, 1); return Data[Offset + offset]; }
        internal uint U32(int offset) { Check(offset, 4); return VisioBinaryData.U32(Data, Offset + offset); }
        internal double Number(int offset) { Check(offset, 8); return VisioBinaryData.Number(Data, Offset + offset); }
        internal string Text(int offset) {
            Check(offset, 0);
            int bytes = Length - offset;
            if ((bytes & 1) != 0) throw new InvalidDataException("Binary Visio UTF-16 text has an odd byte count.");
            return new System.Text.UnicodeEncoding(false, false, true).GetString(Data, Offset + offset, bytes).TrimEnd('\0');
        }
        private void Check(int offset, int length) {
            if (offset < 0 || offset > Length - length) throw new InvalidDataException("Binary Visio field is truncated.");
        }
    }
}
