namespace OfficeIMO.Email.Store;

internal sealed class PstWriterNode {
    internal PstWriterNode(uint nid, uint parentNid, ulong dataBid, ulong subnodeBid = 0) {
        Nid = nid;
        ParentNid = parentNid;
        DataBid = dataBid;
        SubnodeBid = subnodeBid;
    }

    internal uint Nid { get; }
    internal uint ParentNid { get; }
    internal ulong DataBid { get; }
    internal ulong SubnodeBid { get; }
}

internal sealed class PstWriterBlock {
    internal PstWriterBlock(ulong bid, long offset, int length) {
        Bid = bid;
        Offset = offset;
        Length = length;
    }

    internal ulong Bid { get; }
    internal long Offset { get; }
    internal int Length { get; }
}

internal readonly struct PstWriterPageReference {
    internal PstWriterPageReference(ulong key, ulong bid, long offset) {
        Key = key;
        Bid = bid;
        Offset = offset;
    }

    internal ulong Key { get; }
    internal ulong Bid { get; }
    internal long Offset { get; }
}

internal readonly struct PstWriterAllocation {
    internal PstWriterAllocation(long offset, int length) {
        Offset = offset;
        Length = length;
    }

    internal long Offset { get; }
    internal int Length { get; }
}

internal readonly struct PstWriterTreeRoot {
    internal PstWriterTreeRoot(ulong bid, long offset) {
        Bid = bid;
        Offset = offset;
    }

    internal ulong Bid { get; }
    internal long Offset { get; }
}
