namespace OfficeIMO.Publisher.Internal;

/// <summary>Plans utility surgery in one existing native drawing, retaining allocation history and unrelated records.</summary>
internal sealed class PublisherEscherTableBorderPlan {
    internal PublisherEscherShapeEdits Shapes { get; } = new();
    internal List<(PublisherNativeValueSlot Slot, uint Value)> Scalars { get; } = new();

    internal static PublisherEscherTableBorderPlan? Create(byte[] source, PublisherEscherData art,
        IReadOnlyDictionary<uint, IReadOnlyList<PublisherTableBorderSegment>> desired, PublisherParseContext context) {
        var data = new PublisherBinaryData(source, "Escher/EscherStm");
        var reader = new PublisherEscherReader(data, null, context);
        IReadOnlyList<PublisherEscherRecord> top = reader.Records(0, source.Length);
        PublisherEscherRecord[] dggs = top.Where(item => item.Kind == 0xF000).ToArray();
        PublisherEscherRecord[] drawings = top.Where(item => item.Kind == 0xF002).ToArray();
        if (dggs.Length != 1 || drawings.Length == 0) return null;
        var spids = new HashSet<uint>();
        long total = 0;
        PublisherEscherRecord? utility = null, fdg = null, group = null;
        uint drawingId = 0;
        foreach (PublisherEscherRecord drawing in drawings) {
            context.Token.ThrowIfCancellationRequested();
            IReadOnlyList<PublisherEscherRecord> children = reader.Records(drawing.Offset, drawing.End);
            PublisherEscherRecord[] counts = children.Where(item => item.Kind == 0xF008).ToArray();
            if (counts.Length != 1 || counts[0].Length != 8) return null;
            int count = CountShapes(drawing.Offset, drawing.End, 0);
            if (count < 0 || data.U32(counts[0].Offset) != count) return null;
            total += count;
            if (drawing.Offset < 12 || data.U32(drawing.Offset - 12) != 2) continue;
            PublisherEscherRecord[] groups = children.Where(item => item.Kind == 0xF003).ToArray();
            if (utility != null || groups.Length != 1) return null;
            utility = drawing; fdg = counts[0]; group = groups[0]; drawingId = (uint)(fdg.Initial >> 4);
            if (drawingId == 0 || (fdg.Initial & 15) != 0) return null;
        }
        if (utility == null || fdg == null || group == null) return null;
        IReadOnlyList<PublisherEscherRecord> utilityShapes = reader.Records(group.Offset, group.End);
        if (utilityShapes.Any(item => item.Kind != 0xF004) || utilityShapes.Count != data.U32(fdg.Offset)) return null;
        var bodies = new HashSet<int>(utilityShapes.Select(item => item.Offset));
        PublisherEscherRecord[] globals = reader.Records(dggs[0].Offset, dggs[0].End).Where(item => item.Kind == 0xF006).ToArray();
        if (globals.Length != 1 || globals[0].Length < 16) return null;
        PublisherEscherRecord fdgg = globals[0];
        uint clusters = data.U32(fdgg.Offset + 4);
        if (clusters < 2 || 16L + (clusters - 1L) * 8 != fdgg.Length
            || data.U32(fdgg.Offset + 8) != total || data.U32(fdgg.Offset + 12) != drawings.Length) return null;
        uint cursor = data.U32(fdg.Offset + 4), cluster = cursor >> 10;
        if (cluster == 0 || cluster >= clusters) return null;
        int clusterSlot = checked(fdgg.Offset + 16 + ((int)cluster - 1) * 8);
        uint used = data.U32(clusterSlot + 4);
        if (data.U32(clusterSlot) != drawingId || used > 1024 || used != (cursor & 1023) + 1) return null;
        uint maximum = data.U32(fdgg.Offset);
        if (spids.Count == 0 || maximum <= spids.Max()) return null;
        foreach (PublisherEscherRecord shape in utilityShapes) {
            PublisherEscherRecord descriptor = reader.Records(shape.Offset, shape.End).Single(item => item.Kind == 0xF00A);
            uint id = data.U32(descriptor.Offset);
            if (id >> 10 == cluster && id > cursor) return null;
        }
        var plan = new PublisherEscherTableBorderPlan();
        var added = new List<byte[]>();
        int additions = 0;
        foreach (var table in desired) {
            context.Token.ThrowIfCancellationRequested();
            var wanted = table.Value.ToDictionary(Key);
            var existing = new HashSet<(int, int, int, int)>();
            foreach (PublisherTableArtwork owner in art.TableArtwork.Where(item => item.TableId == table.Key && item.Value(0x2001) is 1 or 2)) {
                PublisherNativeShapeBinding? binding = owner.Shape.NativeBinding;
                if (binding == null || !binding.Matches(data) || !bodies.Contains(binding.BodyOffset)) return null;
                var key = ((int)owner.Value(0x2004), (int)owner.Value(0x2005), (int)owner.Value(0x2006), (int)owner.Value(0x2007));
                if (!existing.Add(key)) return null;
                if (!wanted.ContainsKey(key)) plan.Shapes.Removed.Add(binding.BodyOffset);
            }
            foreach (PublisherTableBorderSegment border in table.Value) {
                context.Token.ThrowIfCancellationRequested();
                if (existing.Contains(Key(border))) continue;
                if (used + additions >= 1024) return null; // A new cluster needs independent qualification.
                uint id = (cluster << 10) + used + (uint)additions;
                if (!spids.Add(id)) return null;
                byte[]? bytes = PublisherTableBorderEncoding.Write(table.Key, id, border);
                if (bytes == null) return null;
                added.Add(bytes); additions++;
            }
        }
        int delta = additions - plan.Shapes.Removed.Count;
        long newTotal = total + delta, newDrawingCount = utilityShapes.Count + (long)delta;
        if (newTotal < 0 || newTotal > context.Options.Limits.MaxItems || newDrawingCount < 1 || newDrawingCount > uint.MaxValue) return null;
        if (additions > 0) {
            long length = added.Sum(bytes => (long)bytes.Length);
            if (length > int.MaxValue) return null;
            using var stream = new MemoryStream((int)length);
            foreach (byte[] bytes in added) { context.Token.ThrowIfCancellationRequested(); stream.Write(bytes, 0, bytes.Length); }
            plan.Shapes.Appended.Add(group.Offset, stream.ToArray());
            uint last = (cluster << 10) + used + (uint)additions - 1;
            Scalar(fdg.Offset + 4, last); Scalar(clusterSlot + 4, used + (uint)additions);
            Scalar(fdgg.Offset, Math.Max(maximum, checked(last + 1)));
        }
        Scalar(fdg.Offset, (uint)newDrawingCount); Scalar(fdgg.Offset + 8, (uint)newTotal);
        return plan;

        void Scalar(int offset, uint value) { if (data.U32(offset) != value) plan.Scalars.Add((new(offset, data.U32(offset)), value)); }
        int CountShapes(int start, int end, int depth) {
            context.CheckDepth(depth); int count = 0;
            foreach (PublisherEscherRecord item in reader.Records(start, end)) {
                context.Token.ThrowIfCancellationRequested();
                if (item.Kind == 0xF004) {
                    PublisherEscherRecord[] descriptors = reader.Records(item.Offset, item.End).Where(value => value.Kind == 0xF00A).ToArray();
                    if (descriptors.Length != 1 || descriptors[0].Length != 8 || !spids.Add(data.U32(descriptors[0].Offset))) return -1;
                    // Native FDG/FDGG counts exclude the drawing background,
                    // although its SPID still belongs to the allocation history.
                    if ((data.U32(descriptors[0].Offset + 4) & 0x400) == 0) count++;
                } else if ((item.Initial & 15) == 15 && item.Kind != 0xF001) {
                    int children = CountShapes(item.Offset, item.End, depth + 1);
                    if (children < 0) return -1;
                    count = checked(count + children);
                }
            }
            return count;
        }
    }

    internal static (int, int, int, int) Key(PublisherTableBorderSegment border) => (border.FirstRow, border.FirstColumn, border.LastRow, border.LastColumn);
}
