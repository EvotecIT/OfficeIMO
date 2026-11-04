namespace OfficeIMO.Publisher.Internal;

/// <summary>Replaces admitted typed fields while retaining every unrelated source token and its order.</summary>
internal static class PublisherOplWriter {
    internal static byte[] Rewrite(PublisherBinaryData source, int offset, int length,
        IReadOnlyList<PublisherBlock> fields, IReadOnlyDictionary<byte, byte[]> changes, PublisherParseContext context) {
        var missing = changes.Keys.Where(id => fields.All(field => field.Id != id)).OrderBy(id => id).ToList();
        source.Range(offset, length);
        using var stream = new MemoryStream(length + changes.Values.Sum(value => value.Length));
        using (var writer = new BinaryWriter(stream, System.Text.Encoding.UTF8, leaveOpen: true)) {
            writer.Write(0U);
            foreach (PublisherBlock field in fields) {
                context.Record();
                foreach (byte id in missing.Where(id => id < field.Id).ToArray()) { writer.Write(changes[id]); missing.Remove(id); }
                if (changes.TryGetValue(field.Id, out byte[]? replacement)) writer.Write(replacement);
                else writer.Write(source.Bytes, field.Offset - 2, field.Length + 2);
            }
            foreach (byte id in missing) writer.Write(changes[id]);
            writer.Flush(); stream.Position = 0; writer.Write(checked((uint)stream.Length));
        }
        return stream.ToArray();
    }

    internal static byte[] Boolean(byte id, bool value) => new[] { id, value ? (byte)0x0A : (byte)0x02 };
    internal static byte[] UInt16(byte id, ushort value, byte type = 0x12) => new[] { id, type, (byte)value, (byte)(value >> 8) };
    internal static byte[] UInt32(byte id, uint value, byte type = 0x22) => new[] { id, type, (byte)value, (byte)(value >> 8), (byte)(value >> 16), (byte)(value >> 24) };
    /// <summary>Creates a length-prefixed chunk without an enclosing field identifier or type.</summary>
    internal static byte[] Chunk(params byte[][] fields) {
        byte[] field = Nested(0, 0x88, fields);
        var chunk = new byte[field.Length - 2];
        Buffer.BlockCopy(field, 2, chunk, 0, chunk.Length);
        return chunk;
    }
    internal static byte[] Nested(byte id, byte type, params byte[][] children) {
        int length = checked(4 + children.Sum(child => child.Length)); var bytes = new byte[length + 2]; bytes[0] = id; bytes[1] = type;
        for (int i = 0; i < 4; i++) bytes[2 + i] = (byte)(length >> (i * 8));
        int position = 6;
        foreach (byte[] child in children) { Buffer.BlockCopy(child, 0, bytes, position, child.Length); position += child.Length; }
        return bytes;
    }
}
