namespace OfficeIMO.Publisher.Internal;

internal static class PublisherQuillDirectory {
    internal static IReadOnlyList<PublisherQuillChunk> Read(PublisherBinaryData data, PublisherParseContext context) {
        data.Range(0, 32);
        if (data.Tag(0) != "CHNK" || data.Tag(4) != "INK ") throw new InvalidDataException("Publisher Quill signature is invalid.");
        var result = new List<PublisherQuillChunk>();
        var visited = new HashSet<uint>();
        uint next = 0x18;
        while (next != uint.MaxValue) {
            context.Record();
            if (!visited.Add(next)) throw new InvalidDataException("Cyclic Publisher Quill directory.");
            int directory = data.Offset(next);
            data.Range(directory, 8);
            int count = data.U16(directory + 2);
            data.Range(directory + 8, count * 24);
            for (int i = 0; i < count; i++) {
                context.Record();
                int position = directory + 8 + i * 24;
                if (data.U16(position) != 24) throw new InvalidDataException("Unsupported Publisher Quill descriptor size.");
                int offset = data.Offset(data.U32(position + 16));
                int length = data.Offset(data.U32(position + 20));
                data.Range(offset, length);
                result.Add(new PublisherQuillChunk(data.Tag(position + 2), data.U16(position + 6), offset, length));
            }
            next = data.U32(directory + 4);
        }
        return result;
    }
    internal static PublisherQuillChunk? Single(IReadOnlyList<PublisherQuillChunk> chunks, string tag) {
        PublisherQuillChunk? result = null;
        foreach (PublisherQuillChunk chunk in chunks) {
            if (chunk.Tag != tag) continue;
            if (result != null) throw new InvalidDataException($"Ambiguous Publisher Quill {tag} section.");
            result = chunk;
        }
        return result;
    }
}

internal sealed class PublisherQuillChunk {
    internal PublisherQuillChunk(string tag, int id, int offset, int length) { Tag = tag; Id = id; Offset = offset; Length = length; }
    internal string Tag { get; }
    internal int Id { get; }
    internal int Offset { get; }
    internal int Length { get; }
    internal int End => Offset + Length;
}
