namespace OfficeIMO.Publisher.Internal;

/// <summary>Plans utility surgery in one existing native drawing, retaining allocation history and unrelated records.</summary>
internal sealed class PublisherEscherTableBorderPlan {
    internal PublisherEscherShapeEdits Shapes { get; } = new();
    internal List<(PublisherNativeValueSlot Slot, uint Value)> Scalars { get; } = new();

    internal static PublisherEscherTableBorderPlan? Create(byte[] source, PublisherEscherData art,
        IReadOnlyDictionary<uint, IReadOnlyList<PublisherTableBorderSegment>> desired, PublisherParseContext context,
        IReadOnlyDictionary<uint, PublisherTableState>? newTables = null,
        IReadOnlyDictionary<uint, PublisherTableState>? partialTables = null) {
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
        var globalRecords = reader.Records(dggs[0].Offset, dggs[0].End);
        PublisherEscherRecord[] globals = globalRecords.Where(item => item.Kind == 0xF006).ToArray();
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
        if (newTables?.Count > 0 || partialTables?.Count > 0) {
            if (newTables?.Keys.Any(desired.ContainsKey) == true
                || (partialTables != null && newTables?.Keys.Any(partialTables.ContainsKey) == true) || !art.HasTableCellDefault
                || !QualifiedUtilityDefaults(data, reader, dggs[0], utilityShapes, drawingId, context)) return null;
            var ordinary = new HashSet<int>();
            foreach (var drawing in drawings.Where(item => item.Offset >= 12 && data.U32(item.Offset - 12) == 0))
                foreach (var container in reader.Records(drawing.Offset, drawing.End).Where(item => item.Kind == 0xF003))
                    foreach (var shape in reader.Records(container.Offset, container.End).Where(item => item.Kind == 0xF004)) ordinary.Add(shape.Offset);
            var inputs = new List<PublisherTableUtilityInput>();
            var selectedCells = new List<PublisherTableCellUtilityInput>();
            if (newTables != null) foreach (var item in newTables.OrderBy(item => item.Key)) {
                context.Token.ThrowIfCancellationRequested();
                var table = item.Value;
                if (!QualifiedTable(item.Key, table) || art.TableArtwork.Any(owner => owner.TableId == item.Key)) return null;
                inputs.Add(new(table.Id, table.Rows.Count, table.Columns.Count, table.Width, table.Height, table.CellFills, table.BorderSegments));
            }
            if (partialTables != null) foreach (var item in partialTables.OrderBy(item => item.Key)) {
                context.Token.ThrowIfCancellationRequested();
                var table = item.Value;
                if (!QualifiedTable(item.Key, table) || !TryCellOwners(data, art.TableArtwork.Where(owner => owner.TableId == item.Key),
                    table, bodies, context, out var existingCells)) return null;
                int start = selectedCells.Count;
                for (int index = 0; index < table.CellFills.Count; index++) {
                    context.Token.ThrowIfCancellationRequested();
                    if (!existingCells.Contains(index)) selectedCells.Add(new(table.Id, index, table.CellFills[index]));
                }
                if (selectedCells.Count == start) return null;
            }
            var cells = PublisherTableUtilityDrawingEncoding.WriteCells((cluster << 10) + used, inputs, selectedCells, (int)(1024 - used), context.Token);
            if (cells == null) return null;
            for (uint index = 0; index < cells.Value.Count; index++)
                if (!spids.Add((cluster << 10) + used + index)) return null;
            added.Add(cells.Value.Bytes); additions = (int)cells.Value.Count;

            bool QualifiedTable(uint id, PublisherTableState table) {
                var owners = art.Shapes.Where(shape => shape.Id == id).ToArray();
                return table.Id == id && table.HasPlacement && !table.Hidden && table.BoundsEditLimitation == null
                    && table.Rows.Count > 0 && table.Columns.Count > 0 && (long)table.Rows.Count * table.Columns.Count == table.Cells.Count
                    && !table.Cells.Any(cell => cell.RowSpan != 1 || cell.ColumnSpan != 1)
                    && owners.Length == 1 && owners[0].NativeBinding is { } binding && binding.Matches(data) && ordinary.Contains(binding.BodyOffset);
            }
        }
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
            // Fresh cells and edges inherit Dgg properties directly. A local
            // utility-default value must not mask unsafe document defaults.
            if (!PublisherEscherTableUtilityPlan.QualifiedDefaults(data, globalRecords, context)) return null;
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

    /// <summary>Qualifies the complete retained ownership before filling holes, including unchanged invisible cells.</summary>
    private static bool TryCellOwners(PublisherBinaryData data, IEnumerable<PublisherTableArtwork> records, PublisherTableState table,
        HashSet<int> bodies, PublisherParseContext context, out HashSet<int> cells) {
        cells = new HashSet<int>();
        var edges = new HashSet<(uint, uint, uint, uint)>();
        foreach (PublisherTableArtwork owner in records) {
            context.Token.ThrowIfCancellationRequested();
            var shape = owner.Shape; var style = shape.Style;
            if (shape.NativeBinding is not { } binding || !binding.Matches(data) || !bodies.Contains(binding.BodyOffset)
                || shape.Type != 1 || shape.Flags != 0xA00 || shape.Hidden
                || shape.Transform.RotationDegrees.GetValueOrDefault() != 0 || shape.Transform.FlipHorizontal || shape.Transform.FlipVertical
                || style.ShadowEnabled == true || style.FillType.GetValueOrDefault() != 0 || style.FillOpacity.GetValueOrDefault(1) != 1) return false;
            uint kind = owner.Value(0x2001);
            if (kind == 0) {
                uint index = owner.Value(0x2003);
                if (index >= table.CellFills.Count || !cells.Add((int)index)
                    || owner.Anchor.Keys.Any(key => key is not (0x6802 or 0x2001 or 0x2003))) return false;
            } else if (kind is 1 or 2) {
                uint row = owner.Value(0x2004), column = owner.Value(0x2005), endRow = owner.Value(0x2006), endColumn = owner.Value(0x2007);
                if (row > table.Rows.Count || endRow > table.Rows.Count || column > table.Columns.Count || endColumn > table.Columns.Count
                    || (kind == 1 ? row != endRow || endColumn <= column : column != endColumn || endRow <= row)
                    || !edges.Add((row, column, endRow, endColumn)) || style.FillEnabled == false
                    || owner.Anchor.Keys.Any(key => key is not (0x6802 or 0x2001 or 0x2004 or 0x2005 or 0x2006 or 0x2007))) return false;
            } else return false;
        }
        return true;
    }

    private static bool QualifiedUtilityDefaults(PublisherBinaryData data, PublisherEscherReader reader, PublisherEscherRecord dgg,
        IReadOnlyList<PublisherEscherRecord> shapes, uint drawingId, PublisherParseContext context) {
        if (shapes.Count < 3) return false;
        var root = reader.Records(shapes[0].Offset, shapes[0].End);
        var fsps = root.Where(item => item.Kind == 0xF00A).ToArray();
        var spaces = root.Where(item => item.Kind == 0xF009).ToArray();
        if (root.Count != 2 || fsps.Length != 1 || fsps[0].Length != 8 || fsps[0].Initial != 2
            || data.U32(fsps[0].Offset) != drawingId << 10 || data.U32(fsps[0].Offset + 4) != 5
            || spaces.Length != 1 || spaces[0].Length != 16 || Enumerable.Range(0, 4).Any(axis => data.U32(spaces[0].Offset + axis * 4) != 0)) return false;
        var globals = reader.Records(dgg.Offset, dgg.End); int cells = 0, borders = 0;
        foreach (var shape in shapes.Skip(1)) {
            context.Token.ThrowIfCancellationRequested();
            var records = reader.Records(shape.Offset, shape.End);
            var anchors = records.Where(item => item.Kind == 0xF010).ToArray();
            if (anchors.Length != 1) return false;
            var anchor = reader.ClientValues(anchors[0]);
            if (anchor.ContainsKey(0x6802)) continue;
            var descriptors = records.Where(item => item.Kind == 0xF00A).ToArray();
            if (descriptors.Length != 1 || data.U32(descriptors[0].Offset + 4) != 0xA00
                || records.Any(item => item.Kind is not (0xF00A or 0xF00B or 0xF122 or 0xF010))
                || !PublisherEscherTableUtilityPlan.QualifiedDefaults(data, globals.Concat(records).ToArray(), context)) return false;
            int type = descriptors[0].Initial >> 4;
            if (type == 1 && anchor.Count == 0) cells++;
            else if (type == 20 && anchor.Count == 1 && anchor.TryGetValue(0x2001, out uint kind) && kind == 1) borders++;
            else return false;
        }
        return cells == 1 && borders == 1;
    }
}
