namespace OfficeIMO.Publisher.Internal;

/// <summary>Composes page relationships, shape edits and owned picture filenames with native chunk-directory relocation.</summary>
internal static class PublisherContentsSourceWriter {
    internal static byte[]? Rewrite(byte[] source, PublisherContents contents, IReadOnlyDictionary<uint, uint> alignments,
        IReadOnlyDictionary<uint, string> pictureNames, IReadOnlyDictionary<uint, uint?> masterAssignments,
        PublisherParseContext context, long maximumBytes, IReadOnlyDictionary<uint, bool>? storyHyphenation = null,
        IReadOnlyDictionary<uint, IReadOnlyList<uint>>? frameWraps = null) {
        var data = new PublisherBinaryData(source, "Contents");
        var blocks = new PublisherBlockReader(data, context);
        var edits = new List<PublisherNativeStreamEdit>();
        var growth = new SortedDictionary<int, int>();
        if (storyHyphenation?.Count > 0) {
            PublisherNativeStreamEdit? edit = PublisherStoryHyphenation.Rewrite(data, contents, storyHyphenation, context);
            if (edit == null) return null;
            edits.Add(edit); growth.Add(edit.Offset, edit.Bytes.Length - edit.Length);
        }
        var shapeChanges = new Dictionary<uint, Dictionary<byte, byte[]>>();
        foreach (var pair in alignments) {
            context.Token.ThrowIfCancellationRequested();
            if (!contents.Shapes.TryGetValue(pair.Key, out PublisherSourceShape? shape) || pair.Value > 2) return null;
            PublisherBlock[] prior = shape.Fields.Where(field => field.Id == 0x35).ToArray();
            if (prior.Length > 1 || prior.Any(field => field.Type != 0x20 || field.Value > 2)) return null;
            shapeChanges.Add(pair.Key, new Dictionary<byte, byte[]> { [0x35] = PublisherOplWriter.UInt32(0x35, pair.Value, 0x20) });
        }
        if (frameWraps != null) foreach (var pair in frameWraps) {
            context.Token.ThrowIfCancellationRequested();
            if (!contents.Shapes.TryGetValue(pair.Key, out PublisherSourceShape? shape) || shape.Chunk.Kind != 1
                || !shape.Value(0x27).HasValue || !contents.ObjectPages.TryGetValue(pair.Key, out uint owner)) return null;
            PublisherBlock[] counts = shape.Fields.Where(field => field.Id == 0x46).ToArray();
            PublisherBlock[] lists = shape.Fields.Where(field => field.Id == 0x47).ToArray();
            if (counts.Length > 1 || counts.Any(field => field.Type != 0x20 || field.Value != shape.WrapObjects.Count)
                || lists.Length > 1 || lists.Any(field => field.Type != 0xA0)) return null;
            if (pair.Value.Count > context.Options.Limits.MaxItems) throw new InvalidDataException("Publisher text-wrap item limit exceeded.");
            var seen = new HashSet<uint>(); var references = new byte[pair.Value.Count][];
            for (int index = 0; index < references.Length; index++) {
                context.Record(); uint id = pair.Value[index];
                if (id == pair.Key || !seen.Add(id) || !contents.Shapes.ContainsKey(id)
                    || !contents.ObjectPages.TryGetValue(id, out uint otherOwner) || otherOwner != owner) return null;
                references[index] = PublisherOplWriter.UInt32(0, id, 0x68);
            }
            if (!shapeChanges.TryGetValue(pair.Key, out var changes)) shapeChanges.Add(pair.Key, changes = new());
            changes.Add(0x46, references.Length == 0 ? Array.Empty<byte>() : PublisherOplWriter.UInt32(0x46, checked((uint)references.Length), 0x20));
            changes.Add(0x47, references.Length == 0 ? Array.Empty<byte>() : PublisherOplWriter.Nested(0x47, 0xA0, references));
        }
        int trailer = data.Offset(data.U32(0x1A));
        IReadOnlyList<PublisherBlock>? sections = null, slots = null;
        PublisherBlock directory = default;
        var additions = new List<(uint Id, byte[] Chunk, PublisherBlock? Slot, uint Owner)>();
        if (pictureNames.Count > 0) {
            sections = blocks.Chunk(trailer);
            var directories = sections.Where(field => field.Type == 0x90).ToArray();
            if (directories.Length != 1 || sections.Count(field => field.Id == directories[0].Id) != 1) return null;
            directory = directories[0]; slots = blocks.Children(directory);
            PublisherBlock[] count = sections.Where(field => field.Id == 1).ToArray(), last = sections.Where(field => field.Id == 2).ToArray();
            if (count.Length != 1 || count[0].Type != 0x20 || count[0].Value != slots.Count
                || last.Length != 1 || last[0].Type != 0x20 || last[0].Value + 1 != count[0].Value) return null;
            foreach (var pair in pictureNames.OrderBy(pair => pair.Key)) {
                context.Record();
                if (!contents.Shapes.TryGetValue(pair.Key, out PublisherSourceShape? shape)) return null;
                PublisherBlock[] reference = shape.Fields.Where(field => field.Id == 0x3A).ToArray();
                PublisherBlock[] marker = shape.Fields.Where(field => field.Id == 3).ToArray();
                if (reference.Length > 1 || marker.Length > 1 || marker.Any(field => field.Type != 8)) return null;
                uint id = checked((uint)(slots.Count + additions.Count));
                byte[] name = PublisherOplWriter.Nested(3, 0xC0, System.Text.Encoding.Unicode.GetBytes(pair.Value + "\0"));
                byte[] chunk;
                PublisherBlock? slot = null;
                if (reference.Length == 0) {
                    // Native resave can drop the local filename reference and
                    // embedded marker while retaining old resource chunks.
                    // Own a fresh resource; never guess which orphan to reuse.
                    if (shape.Chunk.Kind != 0x20) return null;
                    chunk = PublisherOplWriter.Chunk(name);
                } else {
                    if (reference[0].Type != 0x70
                        || !contents.Chunks.TryGetValue(reference[0].Value, out PublisherContentChunk? resource)
                        || resource.Kind != 0x66 || resource.Id >= slots.Count) return null;
                    var fields = blocks.Chunk(resource.Offset, resource.End);
                    if (fields.Count != 1 || fields[0].Id != 3 || fields[0].Type != 0xC0) return null;
                    slot = slots[(int)resource.Id];
                    if (slot.Value.Type != 0x88 || slot.Value.Id != 0) return null;
                    var descriptor = blocks.Children(slot.Value);
                    if (descriptor.Any(field => field.Id is not (2 or 4 or 5 or 6 or 0x0B))
                        || descriptor.Count(field => field.Id == 5 && field.Type == 0x68) != 1) return null;
                    chunk = PublisherOplWriter.Rewrite(data, resource.Offset, data.Offset(data.U32(resource.Offset)), fields,
                        new Dictionary<byte, byte[]> { [3] = name }, context);
                }
                additions.Add((id, chunk, slot, pair.Key));
                if (!shapeChanges.TryGetValue(pair.Key, out var changes)) shapeChanges.Add(pair.Key, changes = new());
                changes.Add(0x3A, PublisherOplWriter.UInt32(0x3A, id, 0x70)); changes.Add(3, new byte[] { 3, 8 });
            }
        }
        foreach (var pair in shapeChanges) {
            PublisherSourceShape shape = contents.Shapes[pair.Key];
            int length = data.Offset(data.U32(shape.Chunk.Offset));
            byte[] replacement = PublisherOplWriter.Rewrite(data, shape.Chunk.Offset, length, shape.Fields, pair.Value, context);
            edits.Add(new(shape.Chunk.Offset, length, replacement));
            growth.Add(shape.Chunk.Offset, replacement.Length - length);
        }
        foreach (var pair in masterAssignments) {
            context.Token.ThrowIfCancellationRequested();
            PublisherSourcePage? page = contents.Pages.SingleOrDefault(item => item.Id == pair.Key);
            if (page == null || page.IsMaster || !page.MasterReferenceQualified || !page.DisplayMaster.HasValue
                || !contents.Chunks.TryGetValue(pair.Key, out PublisherContentChunk? chunk) || chunk.Kind != 0x43) return null;
            if (pair.Value.HasValue && !contents.Pages.Any(item => item.IsMaster && item.Id == pair.Value.Value)) return null;
            var fields = blocks.Chunk(chunk.Offset, chunk.End);
            PublisherBlock[] display = fields.Where(field => field.Id == 7).ToArray(), reference = fields.Where(field => field.Id == 0x0D).ToArray();
            if (display.Length > 1 || display.Any(field => field.Type != 8)
                || reference.Length > 1 || reference.Any(field => field.Type != 0x68)) return null;
            var changes = new Dictionary<byte, byte[]> {
                // Native IgnoreMaster=true writes an explicit false display token;
                // enabling formatting restores the qualified omitted default.
                [7] = pair.Value.HasValue ? Array.Empty<byte>() : new byte[] { 7, 8 }
            };
            if (pair.Value.HasValue) changes.Add(0x0D, PublisherOplWriter.UInt32(0x0D, pair.Value.Value, 0x68));
            int length = data.Offset(data.U32(chunk.Offset));
            byte[] replacement = PublisherOplWriter.Rewrite(data, chunk.Offset, length, fields, changes, context);
            edits.Add(new(chunk.Offset, length, replacement));
            growth.Add(chunk.Offset, replacement.Length - length);
        }
        int delta = growth.Values.Sum();
        if (growth.Values.Any(value => value != 0) || additions.Count > 0) {
            sections ??= blocks.Chunk(trailer);
            PublisherBlock[] directories = sections.Where(field => field.Type == 0x90).ToArray();
            if (directories.Length != 1 || data.U32(8) != source.Length) return null;
            var properties = blocks.Chunk(44, data.Offset(data.U32(0x1E)));
            PublisherBlock[] size = properties.Where(field => field.Id == 8).ToArray(), pointer = properties.Where(field => field.Id == 10).ToArray();
            if (size.Length != 1 || size[0].Type != 0x20 || size[0].Value != source.Length
                || pointer.Length != 1 || pointer[0].Type != 0xB8 || pointer[0].Value != trailer) return null;
            int outputLength = checked(source.Length + delta), outputTrailer = checked(trailer + delta);
            directory = directories[0]; slots ??= blocks.Children(directory);
            uint id = 0; int matched = 0;
            var ordered = growth.ToArray(); int index = 0, shift = 0;
            var relocated = new Dictionary<int, int>();
            foreach (PublisherContentChunk chunk in contents.Chunks.Values.OrderBy(chunk => chunk.Offset)) {
                while (index < ordered.Length && ordered[index].Key < chunk.Offset) shift += ordered[index++].Value;
                relocated.Add(chunk.Offset, checked(chunk.Offset + shift));
            }
            var directoryRows = new List<byte[]>();
            foreach (PublisherBlock slot in slots) {
                if (contents.Chunks.TryGetValue(id, out PublisherContentChunk? chunk)) {
                    if (slot.Type != 0x88) return null;
                    PublisherBlock[] offsets = blocks.Children(slot).Where(field => field.Id == 4).ToArray();
                    if (offsets.Length != 1 || offsets[0].Type != 0xB8 || offsets[0].Value != chunk.Offset) return null;
                    if (additions.Count == 0) Put(offsets[0].Offset, checked((uint)relocated[chunk.Offset]));
                    else directoryRows.Add(RewriteSlot(slot, new Dictionary<byte, byte[]> { [4] = PublisherOplWriter.UInt32(4, checked((uint)relocated[chunk.Offset]), 0xB8) }));
                    matched++;
                } else if (additions.Count > 0) {
                    if (slot.Type == 0x88) return null;
                    directoryRows.Add(source.Skip(slot.Offset - 2).Take(slot.Length + 2).ToArray());
                }
                id++;
            }
            if (matched != contents.Chunks.Count) return null;
            if (additions.Count > 0) {
                int next = checked(trailer + delta);
                foreach (var addition in additions) {
                    var fields = new Dictionary<byte, byte[]> {
                        [4] = PublisherOplWriter.UInt32(4, checked((uint)next), 0xB8),
                        [5] = PublisherOplWriter.UInt32(5, addition.Owner, 0x68)
                    };
                    directoryRows.Add(addition.Slot.HasValue ? RewriteSlot(addition.Slot.Value, fields)
                        : PublisherOplWriter.Nested(0, 0x88, PublisherOplWriter.UInt16(2, 0x66, 0x18), fields[4], fields[5]));
                    next = checked(next + addition.Chunk.Length);
                }
                int trailerLength = data.Offset(data.U32(trailer));
                byte[] rewritten = PublisherOplWriter.Rewrite(data, trailer, trailerLength, sections, new Dictionary<byte, byte[]> {
                    [1] = PublisherOplWriter.UInt32(1, checked((uint)(slots.Count + additions.Count)), 0x20),
                    [2] = PublisherOplWriter.UInt32(2, additions[additions.Count - 1].Id, 0x20),
                    [directory.Id] = PublisherOplWriter.Nested(directory.Id, directory.Type, directoryRows.ToArray())
                }, context);
                long combinedLength = (long)next - (trailer + delta) + rewritten.Length;
                if (combinedLength > maximumBytes) throw new InvalidDataException("Publisher native output byte limit exceeded.");
                var bytes = new byte[(int)combinedLength]; int position = 0;
                foreach (var addition in additions) { addition.Chunk.CopyTo(bytes, position); position += addition.Chunk.Length; }
                rewritten.CopyTo(bytes, position); edits.Add(new(trailer, trailerLength, bytes));
                int totalDelta = checked(delta + bytes.Length - trailerLength);
                outputLength = checked(source.Length + totalDelta); outputTrailer = next;
            }
            Put(8, (uint)outputLength); Put(size[0].Offset, (uint)outputLength);
            Put(0x1A, (uint)outputTrailer); Put(pointer[0].Offset, (uint)outputTrailer);
        }
        return PublisherNativeStreamEdit.Apply(source, edits, maximumBytes, context.Token);

        void Put(int offset, uint value) => edits.Add(new(offset, 4, BitConverter.GetBytes(value)));
        byte[] RewriteSlot(PublisherBlock slot, IReadOnlyDictionary<byte, byte[]> changes) {
            byte[] body = PublisherOplWriter.Rewrite(data, slot.Offset, slot.Length, blocks.Children(slot), changes, context);
            return PublisherOplWriter.Nested(slot.Id, slot.Type, body.Skip(4).ToArray());
        }
    }
}
