namespace OfficeIMO.Publisher.Internal;

/// <summary>
/// Recovers ordinal-only native story chains from the publication's frame and
/// story catalogs. Explicit per-frame links remain the projector's primary contract.
/// </summary>
internal static class PublisherTextFrameCatalogReader {
    internal static void Read(PublisherBinaryData data, PublisherContents publication, PublisherParseContext context) {
        var candidates = publication.Shapes.Values.Where(shape => shape.Chunk.Kind != 0x10 && shape.Value(0x27).HasValue)
            .GroupBy(shape => shape.Value(0x27)!.Value)
            .Where(group => group.Count() > 1 && group.All(shape => !HasExplicitLinks(shape)))
            .ToDictionary(group => group.Key, group => group.ToArray());
        if (candidates.Count == 0) return;

        var blocks = new PublisherBlockReader(data, context);
        PublisherContentChunk document = publication.Chunks.Values.Single(chunk => chunk.Kind == 0x44);
        PublisherBlock? catalogReference = Field(blocks.Chunk(document.Offset, document.End), 0x20);
        if (!catalogReference.HasValue) return;
        IReadOnlyList<PublisherBlock> catalog = Chunk(catalogReference.Value, 0x5B, publication, blocks);
        PublisherBlock? framesReference = Field(catalog, 1), storiesReference = Field(catalog, 2);
        if (!framesReference.HasValue || !storiesReference.HasValue) return;

        var bindings = new Dictionary<uint, List<FrameBinding>>();
        foreach (IReadOnlyList<PublisherBlock> entry in Entries(Chunk(framesReference.Value, 0x61, publication, blocks), blocks, context)) {
            uint storyId = Scalar(entry, 1);
            if (!candidates.ContainsKey(storyId)) continue;
            uint ordinal = OptionalScalar(entry, 2) ?? 0;
            PublisherBlock frame = Field(entry, 3) ?? throw new InvalidDataException("Publisher frame catalog entry has no object reference.");
            if (frame.Type is not (0x68 or 0x70)) throw new InvalidDataException("Invalid Publisher frame catalog object reference encoding.");
            if (!bindings.TryGetValue(storyId, out List<FrameBinding>? values))
                bindings.Add(storyId, values = new List<FrameBinding>());
            values.Add(new FrameBinding(frame.Value, ordinal));
        }
        var stories = new HashSet<uint>();
        foreach (IReadOnlyList<PublisherBlock> entry in Entries(Chunk(storiesReference.Value, 0x65, publication, blocks), blocks, context)) {
            uint storyId = Scalar(entry, 1);
            if (!candidates.ContainsKey(storyId)) continue;
            // Field 7 is IdLayout metadata, independent of the frame bindings.
            // Cardinality comes from the complete frame catalog, not this value.
            if (!stories.Add(storyId)) throw new InvalidDataException("Duplicate Publisher story catalog definition.");
        }
        foreach (var candidate in candidates) {
            context.Record();
            if (!bindings.TryGetValue(candidate.Key, out List<FrameBinding>? mapped) || !stories.Contains(candidate.Key))
                continue;
            PublisherSourceShape[] frames = candidate.Value;
            if (mapped.Count != frames.Length)
                throw new InvalidDataException("Publisher frame catalog does not cover all source story frames.");
            var byId = frames.ToDictionary(frame => frame.Chunk.Id);
            var ordered = new uint[frames.Length];
            var seenIds = new HashSet<uint>();
            var seenOrdinals = new HashSet<uint>();
            foreach (FrameBinding binding in mapped) {
                context.Record();
                if (binding.Ordinal >= ordered.Length || !seenOrdinals.Add(binding.Ordinal) || !seenIds.Add(binding.Id)
                    || !byId.TryGetValue(binding.Id, out PublisherSourceShape? frame)
                    || (frame.Value(0x28) ?? 0) != binding.Ordinal)
                    throw new InvalidDataException("Publisher frame catalog has invalid, duplicate or inconsistent story bindings.");
                ordered[binding.Ordinal] = binding.Id;
            }
            publication.CatalogFrameChains.Add(candidate.Key, ordered);
        }
    }

    private static bool HasExplicitLinks(PublisherSourceShape shape) =>
        Field(shape.Fields, 0x36).HasValue || Field(shape.Fields, 0x37).HasValue;

    private static IReadOnlyList<PublisherBlock> Chunk(PublisherBlock reference, uint kind, PublisherContents publication, PublisherBlockReader blocks) {
        if (reference.Type is not (0x68 or 0x70) || !publication.Chunks.TryGetValue(reference.Value, out PublisherContentChunk? chunk) || chunk.Kind != kind)
            throw new InvalidDataException("Publisher text catalog reference has an invalid record kind or target.");
        return blocks.Chunk(chunk.Offset, chunk.End);
    }

    private static IReadOnlyList<IReadOnlyList<PublisherBlock>> Entries(IReadOnlyList<PublisherBlock> fields, PublisherBlockReader blocks, PublisherParseContext context) {
        uint count = Scalar(fields, 1);
        if (count > context.Options.Limits.MaxItems) throw new InvalidDataException("Publisher text catalog exceeds the item limit.");
        PublisherBlock? array = Field(fields, 2);
        if (!array.HasValue) {
            if (count != 0) throw new InvalidDataException("Publisher text catalog count has no entry array.");
            return Array.Empty<IReadOnlyList<PublisherBlock>>();
        }
        if (array.Value.Type != 0xA0) throw new InvalidDataException("Invalid Publisher text catalog array encoding.");
        IReadOnlyList<PublisherBlock> entries = blocks.Children(array.Value);
        if (entries.Count != count) throw new InvalidDataException("Publisher text catalog count disagrees with its entry array.");
        var result = new List<IReadOnlyList<PublisherBlock>>(entries.Count);
        foreach (PublisherBlock entry in entries) {
            context.Record();
            if (entry.Type != 0x88) throw new InvalidDataException("Invalid Publisher text catalog entry encoding.");
            result.Add(blocks.Children(entry));
        }
        return result;
    }

    private static uint Scalar(IReadOnlyList<PublisherBlock> fields, byte id) => OptionalScalar(fields, id)
        ?? throw new InvalidDataException($"Publisher text catalog field 0x{id:X2} is missing.");

    private static uint? OptionalScalar(IReadOnlyList<PublisherBlock> fields, byte id) {
        PublisherBlock? field = Field(fields, id);
        if (!field.HasValue) return null;
        if (field.Value.Type is not (0x18 or 0x20)) throw new InvalidDataException("Invalid Publisher text catalog scalar encoding.");
        return field.Value.Value;
    }

    private static PublisherBlock? Field(IReadOnlyList<PublisherBlock> fields, byte id) => PublisherContentsReader.Field(fields, id);
    private readonly record struct FrameBinding(uint Id, uint Ordinal);
}

internal sealed partial class PublisherContents {
    /// <summary>Complete native frame-catalog chains for stories without explicit per-frame link records.</summary>
    internal Dictionary<uint, uint[]> CatalogFrameChains { get; } = new();
}
