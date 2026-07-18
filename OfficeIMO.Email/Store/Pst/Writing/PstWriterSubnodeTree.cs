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
        return WriteSorted(file, ordered, ordered.Length);
    }

    internal static ulong WriteSorted(PstWriterFile file,
        IEnumerable<PstWriterSubnode> sortedSubnodes, int count) {
        if (file == null) throw new ArgumentNullException(nameof(file));
        if (count == 0) return 0;

        int leafCapacity = (MaximumBlockPayload - HeaderSize) / LeafEntrySize;
        var level = new List<SubnodeBlockReference>();
        using (IEnumerator<PstWriterSubnode> values = sortedSubnodes.GetEnumerator()) {
            int remaining = count;
            while (remaining > 0) {
                int leafCount = Math.Min(leafCapacity, remaining);
                var payload = new byte[HeaderSize + leafCount * LeafEntrySize];
                payload[0] = 0x02;
                payload[1] = 0x00;
                PstBinary.WriteUInt16(payload, 2, leafCount);
                uint firstNid = 0;
                for (int index = 0; index < leafCount; index++) {
                    if (!values.MoveNext()) {
                        throw new InvalidDataException("The subnode source ended before its declared count.");
                    }
                    PstWriterSubnode subnode = values.Current;
                    if (index == 0) firstNid = subnode.Nid;
                    int entryOffset = HeaderSize + index * LeafEntrySize;
                    PstBinary.WriteUInt32(payload, entryOffset, subnode.Nid);
                    PstBinary.WriteUInt64(payload, entryOffset + 8, subnode.DataBid);
                    PstBinary.WriteUInt64(payload, entryOffset + 16, subnode.SubnodeBid);
                }
                level.Add(new SubnodeBlockReference(firstNid, file.WriteInternalBlock(payload)));
                remaining -= leafCount;
            }
            if (values.MoveNext()) {
                throw new InvalidDataException("The subnode source exceeded its declared count.");
            }
        }

        int indexCapacity = (MaximumBlockPayload - HeaderSize) / IndexEntrySize;
        while (level.Count > 1) {
            var parent = new List<SubnodeBlockReference>();
            for (int offset = 0; offset < level.Count; offset += indexCapacity) {
                int childCount = Math.Min(indexCapacity, level.Count - offset);
                var payload = new byte[HeaderSize + childCount * IndexEntrySize];
                payload[0] = 0x02;
                payload[1] = 0x01;
                PstBinary.WriteUInt16(payload, 2, childCount);
                for (int index = 0; index < childCount; index++) {
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
