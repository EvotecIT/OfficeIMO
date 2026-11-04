namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherModelImporter {
    private static IReadOnlyDictionary<uint, IReadOnlyList<uint>> ImportFrameChains(PublisherContents source, PublisherQuillText text,
        IReadOnlyDictionary<uint, PublisherFrameState> frames, PublisherParseContext context) {
        var result = new Dictionary<uint, IReadOnlyList<uint>>();
        foreach (var group in source.Shapes.Values.Where(shape => shape.Chunk.Kind != 0x10 && shape.Value(0x27).HasValue).GroupBy(shape => shape.Value(0x27)!.Value)) {
            context.Record();
            if (!text.Stories.ContainsKey(group.Key)) continue;
            PublisherSourceShape[] native = group.ToArray();
            var byId = native.ToDictionary(frame => frame.Chunk.Id);
            foreach (PublisherSourceShape frame in native) {
                uint? previous = ImportedFrameLink(frame, 0x36), next = ImportedFrameLink(frame, 0x37);
                if (previous.HasValue && (!byId.TryGetValue(previous.Value, out PublisherSourceShape? before) || ImportedFrameLink(before, 0x37) != frame.Chunk.Id))
                    throw new InvalidDataException("Publisher previous text-frame link is unresolved or not reciprocal.");
                if (next.HasValue && (!byId.TryGetValue(next.Value, out PublisherSourceShape? after) || ImportedFrameLink(after, 0x36) != frame.Chunk.Id))
                    throw new InvalidDataException("Publisher next text-frame link is unresolved or not reciprocal.");
            }
            uint[] ordered;
            if (source.CatalogFrameChains.TryGetValue(group.Key, out uint[]? catalog)) ordered = catalog.ToArray();
            else {
                PublisherSourceShape[] roots = native.Where(frame => !ImportedFrameLink(frame, 0x36).HasValue).ToArray();
                if (roots.Length == 0) throw new InvalidDataException("Cyclic Publisher text-frame chain.");
                if (roots.Length != 1) {
                    context.Add("PUB_TEXT_CHAIN_UNRESOLVED", "A source story has multiple unlinked roots. Frame order is ambiguous, so its text is retained without guessed placement.",
                        OfficeConversionLossKind.Omission, "Quill/story/" + group.Key);
                    result.Add(group.Key, Array.Empty<uint>()); continue;
                }
                var values = new List<uint>(); var visited = new HashSet<uint>();
                PublisherSourceShape? current = roots[0];
                while (current != null) {
                    context.Record();
                    if (!visited.Add(current.Chunk.Id)) throw new InvalidDataException("Cyclic Publisher text-frame chain.");
                    if ((current.Value(0x28) ?? 0) != values.Count) throw new InvalidDataException("Publisher text-frame ordinal disagrees with its native links.");
                    values.Add(current.Chunk.Id);
                    uint? next = ImportedFrameLink(current, 0x37); current = next.HasValue ? byId[next.Value] : null;
                }
                if (values.Count != native.Length) throw new InvalidDataException("Publisher text-frame story contains a disconnected cycle.");
                ordered = values.ToArray();
            }
            foreach (uint id in ordered) {
                PublisherFrameState frame = frames[id];
                if (frame.Available && (frame.ColumnCount == 0 || frame.ColumnCount > 256 || frame.ColumnCount > context.Options.Limits.MaxItems))
                    throw new InvalidDataException("Publisher text column count is outside the supported resource bounds.");
            }
            result.Add(group.Key, Array.AsReadOnly(ordered));
        }
        return result;
    }

    private static uint? ImportedFrameLink(PublisherSourceShape frame, byte id) {
        PublisherBlock? block = PublisherContentsReader.Field(frame.Fields, id);
        if (!block.HasValue) return null;
        if (block.Value.Type is not (0x68 or 0x70)) throw new InvalidDataException("Invalid Publisher text-frame link encoding.");
        return block.Value.Value == 0 ? null : block.Value.Value;
    }
}
