namespace OfficeIMO.Publisher.Internal;

/// <summary>Edits shape vertical alignment, relocating native chunk-directory offsets only when an OPL grows.</summary>
internal static class PublisherContentsLayoutWriter {
    internal static byte[]? Rewrite(byte[] source, PublisherContents contents, IReadOnlyDictionary<uint, uint> alignments,
        PublisherParseContext context, long maximumBytes) {
        var data = new PublisherBinaryData(source, "Contents");
        var blocks = new PublisherBlockReader(data, context);
        var edits = new List<PublisherNativeStreamEdit>();
        var growth = new SortedDictionary<int, int>();
        foreach (var pair in alignments) {
            context.Token.ThrowIfCancellationRequested();
            if (!contents.Shapes.TryGetValue(pair.Key, out PublisherSourceShape? shape) || pair.Value > 2) return null;
            PublisherBlock[] prior = shape.Fields.Where(field => field.Id == 0x35).ToArray();
            if (prior.Length > 1 || prior.Any(field => field.Type != 0x20 || field.Value > 2)) return null;
            int length = data.Offset(data.U32(shape.Chunk.Offset));
            byte[] replacement = PublisherOplWriter.Rewrite(data, shape.Chunk.Offset, length, shape.Fields,
                new Dictionary<byte, byte[]> { [0x35] = PublisherOplWriter.UInt32(0x35, pair.Value, 0x20) }, context);
            edits.Add(new(shape.Chunk.Offset, length, replacement));
            growth.Add(shape.Chunk.Offset, replacement.Length - length);
        }
        int delta = growth.Values.Sum();
        if (delta != 0) {
            int trailer = data.Offset(data.U32(0x1A));
            var sections = blocks.Chunk(trailer);
            PublisherBlock[] directories = sections.Where(field => field.Type == 0x90).ToArray();
            if (directories.Length != 1 || data.U32(8) != source.Length) return null;
            var properties = blocks.Chunk(44, data.Offset(data.U32(0x1E)));
            PublisherBlock[] size = properties.Where(field => field.Id == 8).ToArray(), pointer = properties.Where(field => field.Id == 10).ToArray();
            if (size.Length != 1 || size[0].Type != 0x20 || size[0].Value != source.Length
                || pointer.Length != 1 || pointer[0].Type != 0xB8 || pointer[0].Value != trailer) return null;
            Put(8, checked((uint)(source.Length + delta))); Put(size[0].Offset, checked((uint)(source.Length + delta)));
            Put(0x1A, checked((uint)(trailer + delta))); Put(pointer[0].Offset, checked((uint)(trailer + delta)));
            uint id = 0; int matched = 0;
            var ordered = growth.ToArray(); int index = 0, shift = 0;
            var relocated = new Dictionary<int, int>();
            foreach (PublisherContentChunk chunk in contents.Chunks.Values.OrderBy(chunk => chunk.Offset)) {
                while (index < ordered.Length && ordered[index].Key < chunk.Offset) shift += ordered[index++].Value;
                relocated.Add(chunk.Offset, checked(chunk.Offset + shift));
            }
            foreach (PublisherBlock slot in blocks.Children(directories[0])) {
                if (contents.Chunks.TryGetValue(id, out PublisherContentChunk? chunk)) {
                    if (slot.Type != 0x88) return null;
                    PublisherBlock[] offsets = blocks.Children(slot).Where(field => field.Id == 4).ToArray();
                    if (offsets.Length != 1 || offsets[0].Type != 0xB8 || offsets[0].Value != chunk.Offset) return null;
                    Put(offsets[0].Offset, checked((uint)relocated[chunk.Offset])); matched++;
                }
                id++;
            }
            if (matched != contents.Chunks.Count) return null;
        }
        return PublisherNativeStreamEdit.Apply(source, edits, maximumBytes, context.Token);

        void Put(int offset, uint value) => edits.Add(new(offset, 4, BitConverter.GetBytes(value)));
    }
}
