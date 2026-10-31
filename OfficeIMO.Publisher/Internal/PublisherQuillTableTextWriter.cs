namespace OfficeIMO.Publisher.Internal;

/// <summary>Retains fixed native cell partitions while remapping story-relative TCD endpoints.</summary>
internal static class PublisherQuillTableTextWriter {
    internal static bool TryPlan(PublisherQuillSourceLayout layout,
        IReadOnlyDictionary<uint, PublisherTextStory> before, IReadOnlyDictionary<uint, PublisherTextStory> after,
        IReadOnlyDictionary<uint, PublisherTableStory> beforeTables, IReadOnlyDictionary<uint, PublisherTableStory> afterTables,
        CancellationToken token, out IReadOnlyDictionary<PublisherQuillChunk, byte[]> replacements) {
        var result = new Dictionary<PublisherQuillChunk, byte[]>(); replacements = result;
        PublisherQuillChunk[] chunks = layout.Chunks.Where(chunk => chunk.Tag == "TCD ").ToArray();
        if (chunks.Length != beforeTables.Count || beforeTables.Count != afterTables.Count
            || beforeTables.Keys.Any(id => !afterTables.ContainsKey(id))) return false;
        var seen = new HashSet<uint>(); PublisherBinaryData data = layout.Data;
        foreach (PublisherQuillChunk chunk in chunks) {
            token.ThrowIfCancellationRequested();
            int descriptor = layout.DescriptorOffsets[chunk];
            if (data.U32(descriptor + 8) != 1 || data.Tag(descriptor + 12) != "PLC ") return false;
            int ordinal = chunk.Id;
            if (ordinal < 0 || ordinal >= before.Count) return false;
            uint id = data.U32(layout.StoryIds.Offset + 8 + ordinal * 4);
            if (!seen.Add(id) || !beforeTables.TryGetValue(id, out PublisherTableStory? prior)
                || !afterTables.TryGetValue(id, out PublisherTableStory? current)
                || !before.TryGetValue(id, out PublisherTextStory? priorStory) || !after.TryGetValue(id, out PublisherTextStory? currentStory)
                || !prior.Matches(priorStory) || !current.Matches(currentStory)
                || prior.Segments.Count == 0 || prior.Segments.Count != current.Segments.Count
                || chunk.Length != 12L + prior.Segments.Count * 4L
                || data.U32(chunk.Offset) != prior.Segments.Count - 1 || data.U32(chunk.Offset + 4) != 0
                || data.U32(chunk.Offset + 8) != 0xFF00) return false;
            byte[] bytes = new byte[chunk.Length]; Buffer.BlockCopy(data.Bytes, chunk.Offset, bytes, 0, bytes.Length);
            int oldEnd = 0, newEnd = 0;
            for (int cell = 0; cell < prior.Segments.Count; cell++) {
                token.ThrowIfCancellationRequested(); PublisherTextStory a = prior.Segments[cell], b = current.Segments[cell];
                if (a.Id != id || b.Id != id || !a.HasFinalParagraphTerminator || !b.HasFinalParagraphTerminator
                    || a.Paragraphs.Count != b.Paragraphs.Count) return false;
                oldEnd = checked(oldEnd + a.Text.Length); newEnd = checked(newEnd + b.Text.Length);
                int adjustment = cell == prior.Segments.Count - 1 ? 0 : 1;
                if (data.U32(chunk.Offset + 12 + cell * 4) != oldEnd - adjustment) return false;
                BitConverter.GetBytes(checked((uint)(newEnd - adjustment))).CopyTo(bytes, 12 + cell * 4);
            }
            if (oldEnd != priorStory.Text.Length || newEnd != currentStory.Text.Length) return false;
            if (!bytes.SequenceEqual(data.Bytes.Skip(chunk.Offset).Take(chunk.Length))) result.Add(chunk, bytes);
        }
        return true;
    }
}
