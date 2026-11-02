namespace OfficeIMO.Publisher.Internal;

/// <summary>Initializes table utility records in compact and native-normalized empty-drawing source profiles.</summary>
internal static partial class PublisherEscherTableUtilityPlan {
    internal static PublisherEscherShapeEdits? Create(byte[] source, PublisherEscherData art,
        IReadOnlyDictionary<uint, PublisherTableState> tables, PublisherParseContext context,
        out IReadOnlyList<(PublisherNativeValueSlot Slot, uint Value)> scalars) {
        scalars = Array.Empty<(PublisherNativeValueSlot, uint)>();
        if (tables.Count == 0 || art.TableArtwork.Count != 0 || art.HasTableCellDefault) return null;
        var data = new PublisherBinaryData(source, "Escher/EscherStm");
        var reader = new PublisherEscherReader(data, null, context);
        SourceProfile? profile = ReadProfile(data, reader, context);
        if (profile == null) return null;
        var inputs = new List<PublisherTableUtilityInput>();
        foreach (var entry in tables.OrderBy(item => item.Key)) {
            context.Token.ThrowIfCancellationRequested();
            PublisherTableState table = entry.Value;
            PublisherEscherShape[] owners = art.Shapes.Where(item => item.Id == entry.Key).ToArray();
            if (!table.HasPlacement || table.Hidden || table.BoundsEditLimitation != null
                || table.Cells.Any(cell => cell.RowSpan != 1 || cell.ColumnSpan != 1)
                || owners.Length != 1 || owners[0].NativeBinding is not { } binding || !binding.Matches(data)
                || !profile.OrdinaryBodies.Contains(binding.BodyOffset)) return null;
            inputs.Add(new(table.Id, table.Rows.Count, table.Columns.Count, table.Width, table.Height, table.CellFills, table.BorderSegments));
        }
        int available = (int)Math.Min(1024L, context.Options.Limits.MaxItems - (long)profile.OrdinaryCount - 1);
        PublisherTableUtilityDrawing? utility = PublisherTableUtilityDrawingEncoding.Write(3, inputs, available, context.Token);
        if (utility == null) return null;
        byte[] replacement = new byte[48];
        Buffer.BlockCopy(source, profile.Allocation.Offset - 8, replacement, 0, 32);
        Write(4, 40); Write(8, utility.LastSpid + 1); Write(12, 4);
        Write(16, (uint)profile.OrdinaryCount + utility.Count + 1); Write(20, 3);
        Write(32, 2); Write(36, profile.AuxiliaryUsed); Write(40, 3); Write(44, utility.Count);
        var edits = new PublisherEscherShapeEdits {
            AllocationRecord = new(profile.Allocation.Offset - 8, profile.Allocation.Length + 8, replacement)
        };
        if (profile.EmptyUtility != null) {
            // Native resave can introduce root-only drawing scopes without
            // allocating their IDCL entries. Retain both roots/background and
            // materialize the known allocation slots only when adding utilities.
            var fresh = new PublisherEscherReader(new PublisherBinaryData(utility.Bytes, "new table drawing"), null, context);
            var newDrawing = fresh.Records(0, utility.Bytes.Length).Single();
            var group = fresh.Records(newDrawing.Offset, newDrawing.End).Single(item => item.Kind == 0xF003);
            var root = fresh.Records(group.Offset, group.End)[0];
            edits.Appended.Add(profile.EmptyUtility.Group.Offset, utility.Bytes.Skip(root.End).Take(group.End - root.End).ToArray());
            scalars = new (PublisherNativeValueSlot Slot, uint Value)[] {
                (new(profile.EmptyUtility.Fdg.Offset, data.U32(profile.EmptyUtility.Fdg.Offset)), utility.Count),
                (new(profile.EmptyUtility.Fdg.Offset + 4, data.U32(profile.EmptyUtility.Fdg.Offset + 4)), utility.LastSpid)
            };
        } else {
            using var inserted = new MemoryStream();
            using (var writer = new BinaryWriter(inserted, System.Text.Encoding.UTF8, leaveOpen: true)) {
                writer.Write(2U); writer.Write(utility.Bytes);
                writer.Write(1U); writer.Write(PublisherTableUtilityDrawingEncoding.EmptyDrawing(2));
            }
            edits = new PublisherEscherShapeEdits {
                AllocationRecord = edits.AllocationRecord,
                DrawingInsertion = new(profile.Ordinary.Offset - 12, 0, inserted.ToArray())
            };
        }
        return edits;

        void Write(int offset, uint value) => BitConverter.GetBytes(value).CopyTo(replacement, offset);
    }

    private sealed record SourceProfile(PublisherEscherRecord Allocation, PublisherEscherRecord Ordinary,
        HashSet<int> OrdinaryBodies, int OrdinaryCount, uint AuxiliaryUsed, DrawingProfile? EmptyUtility);
    private sealed record DrawingProfile(PublisherEscherRecord Fdg, PublisherEscherRecord Group,
        IReadOnlyList<PublisherEscherRecord> Shapes, uint Cursor);
}
