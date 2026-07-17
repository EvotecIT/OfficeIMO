namespace OfficeIMO.Email.Store;

internal sealed class PstWriterSubnode {
    internal PstWriterSubnode(uint nid, ulong dataBid, ulong subnodeBid = 0) {
        Nid = nid;
        DataBid = dataBid;
        SubnodeBid = subnodeBid;
    }

    internal uint Nid { get; }
    internal ulong DataBid { get; }
    internal ulong SubnodeBid { get; }
}

internal static class PstWriterSubnodeTree {
    private const int MaximumBlockPayload = 8176;
    private const int HeaderSize = 8;
    private const int LeafEntrySize = 24;
    private const int IndexEntrySize = 16;

    internal static ulong Write(PstWriterFile file, IEnumerable<PstWriterSubnode> subnodes) {
        if (file == null) throw new ArgumentNullException(nameof(file));
        PstWriterSubnode[] ordered = subnodes
            .GroupBy(item => item.Nid)
            .Select(group => group.Last())
            .OrderBy(item => item.Nid)
            .ToArray();
        if (ordered.Length == 0) return 0;

        int leafCapacity = (MaximumBlockPayload - HeaderSize) / LeafEntrySize;
        var level = new List<SubnodeBlockReference>();
        for (int offset = 0; offset < ordered.Length; offset += leafCapacity) {
            int count = Math.Min(leafCapacity, ordered.Length - offset);
            var payload = new byte[HeaderSize + count * LeafEntrySize];
            payload[0] = 0x02;
            payload[1] = 0x00;
            PstBinary.WriteUInt16(payload, 2, count);
            for (int index = 0; index < count; index++) {
                PstWriterSubnode subnode = ordered[offset + index];
                int entryOffset = HeaderSize + index * LeafEntrySize;
                PstBinary.WriteUInt32(payload, entryOffset, subnode.Nid);
                PstBinary.WriteUInt64(payload, entryOffset + 8, subnode.DataBid);
                PstBinary.WriteUInt64(payload, entryOffset + 16, subnode.SubnodeBid);
            }
            level.Add(new SubnodeBlockReference(
                ordered[offset].Nid, file.WriteInternalBlock(payload)));
        }

        int indexCapacity = (MaximumBlockPayload - HeaderSize) / IndexEntrySize;
        while (level.Count > 1) {
            var parent = new List<SubnodeBlockReference>();
            for (int offset = 0; offset < level.Count; offset += indexCapacity) {
                int count = Math.Min(indexCapacity, level.Count - offset);
                var payload = new byte[HeaderSize + count * IndexEntrySize];
                payload[0] = 0x02;
                payload[1] = 0x01;
                PstBinary.WriteUInt16(payload, 2, count);
                for (int index = 0; index < count; index++) {
                    SubnodeBlockReference child = level[offset + index];
                    int entryOffset = HeaderSize + index * IndexEntrySize;
                    PstBinary.WriteUInt32(payload, entryOffset, child.FirstNid);
                    PstBinary.WriteUInt64(payload, entryOffset + 8, child.Bid);
                }
                parent.Add(new SubnodeBlockReference(
                    level[offset].FirstNid, file.WriteInternalBlock(payload)));
            }
            level = parent;
        }
        return level[0].Bid;
    }

    private readonly struct SubnodeBlockReference {
        internal SubnodeBlockReference(uint firstNid, ulong bid) {
            FirstNid = firstNid;
            Bid = bid;
        }

        internal uint FirstNid { get; }
        internal ulong Bid { get; }
    }
}
