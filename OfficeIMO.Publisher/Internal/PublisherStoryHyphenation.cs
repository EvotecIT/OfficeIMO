namespace OfficeIMO.Publisher.Internal;

/// <summary>Reads and rewrites the qualified Contents story flag independently of Quill text and paragraph formatting.</summary>
internal static class PublisherStoryHyphenation {
    internal static void Read(PublisherBinaryData data, PublisherContents publication, PublisherParseContext context) {
        var catalog = Find(data, publication, context);
        if (catalog == null) return;
        var blocks = new PublisherBlockReader(data, context);
        foreach (PublisherBlock entry in catalog.Entries) {
            context.Record();
            var fields = blocks.Children(entry);
            uint id = StoryId(fields);
            bool? enabled = Setting(fields);
            if (publication.StoryHyphenation.ContainsKey(id)) throw new InvalidDataException("Duplicate Publisher story catalog definition.");
            publication.StoryHyphenation.Add(id, enabled);
            if (!enabled.HasValue) context.Add("PUB_HYPHENATION_SETTING_UNASSESSED",
                "The native story hyphenation flag is outside the qualified boolean profile and remains opaque.",
                OfficeConversionLossKind.Unassessed, "Contents/story/" + id);
        }
    }

    internal static PublisherNativeStreamEdit? Rewrite(PublisherBinaryData data, PublisherContents publication,
        IReadOnlyDictionary<uint, bool> changes, PublisherParseContext context) {
        var catalog = Find(data, publication, context);
        if (catalog == null) return null;
        var blocks = new PublisherBlockReader(data, context);
        var entries = new List<byte[]>(); var seen = new HashSet<uint>(); var changed = new HashSet<uint>();
        foreach (PublisherBlock entry in catalog.Entries) {
            context.Record();
            var fields = blocks.Children(entry); uint id = StoryId(fields);
            if (!seen.Add(id)) return null;
            if (changes.TryGetValue(id, out bool enabled)) {
                if (!Setting(fields).HasValue) return null;
                byte[] body = PublisherOplWriter.Rewrite(data, entry.Offset, entry.Length, fields,
                    new Dictionary<byte, byte[]> { [4] = new byte[] { 4, enabled ? (byte)8 : (byte)0 } }, context);
                entries.Add(PublisherOplWriter.Nested(entry.Id, entry.Type, body.Skip(4).ToArray()));
                changed.Add(id);
            } else entries.Add(data.Bytes.AsSpan(entry.Offset - 2, entry.Length + 2).ToArray());
        }
        if (changed.Count != changes.Count) return null;
        byte[] replacement = PublisherOplWriter.Rewrite(data, catalog.Chunk.Offset, data.Offset(data.U32(catalog.Chunk.Offset)),
            catalog.Fields, new Dictionary<byte, byte[]> {
                [2] = PublisherOplWriter.Nested(catalog.Array.Id, catalog.Array.Type, entries.ToArray())
            }, context);
        return new PublisherNativeStreamEdit(catalog.Chunk.Offset, data.Offset(data.U32(catalog.Chunk.Offset)), replacement);
    }

    // Null is deliberately different from an omitted flag inside a valid entry:
    // only that omission has the independently observed enabled default.
    private static Catalog? Find(PublisherBinaryData data, PublisherContents publication, PublisherParseContext context) {
        var blocks = new PublisherBlockReader(data, context);
        var document = publication.Chunks.Values.Single(chunk => chunk.Kind == 0x44);
        var reference = PublisherContentsReader.Field(blocks.Chunk(document.Offset, document.End), 0x20);
        if (!reference.HasValue) return null;
        var textCatalog = Resolve(reference.Value, 0x5B);
        var storyReference = PublisherContentsReader.Field(blocks.Chunk(textCatalog.Offset, textCatalog.End), 2);
        if (!storyReference.HasValue) return null;
        var chunk = Resolve(storyReference.Value, 0x65);
        var fields = blocks.Chunk(chunk.Offset, chunk.End);
        var count = PublisherContentsReader.Field(fields, 1) ?? throw new InvalidDataException("Publisher story catalog count is missing.");
        if (count.Type is not (0x18 or 0x20) || count.Value > context.Options.Limits.MaxItems)
            throw new InvalidDataException("Publisher story catalog count is invalid or exceeds the item limit.");
        var array = PublisherContentsReader.Field(fields, 2);
        if (!array.HasValue && count.Value == 0) return null;
        if (!array.HasValue || array.Value.Type != 0xA0) throw new InvalidDataException("Invalid Publisher story catalog array.");
        var entries = blocks.Children(array.Value);
        if (entries.Count != count.Value || entries.Any(entry => entry.Type != 0x88))
            throw new InvalidDataException("Publisher story catalog entries disagree with their count or encoding.");
        return new Catalog(chunk, fields, array.Value, entries);

        PublisherContentChunk Resolve(PublisherBlock value, uint kind) {
            if (value.Type is not (0x68 or 0x70) || !publication.Chunks.TryGetValue(value.Value, out PublisherContentChunk? target) || target.Kind != kind)
                throw new InvalidDataException("Invalid Publisher story catalog reference.");
            return target;
        }
    }

    private static uint StoryId(IReadOnlyList<PublisherBlock> fields) {
        var id = PublisherContentsReader.Field(fields, 1);
        if (!id.HasValue || id.Value.Type is not (0x18 or 0x20)) throw new InvalidDataException("Invalid Publisher catalog story identifier.");
        return id.Value.Value;
    }

    private static bool? Setting(IReadOnlyList<PublisherBlock> fields) {
        var flag = PublisherContentsReader.Field(fields, 4);
        return !flag.HasValue || flag.Value.Type == 8 ? true : flag.Value.Type == 0 ? false : null;
    }

    private sealed record Catalog(PublisherContentChunk Chunk, IReadOnlyList<PublisherBlock> Fields,
        PublisherBlock Array, IReadOnlyList<PublisherBlock> Entries);
}

internal sealed partial class PublisherContents {
    internal Dictionary<uint, bool?> StoryHyphenation { get; } = new();
}
