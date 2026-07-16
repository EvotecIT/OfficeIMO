namespace OfficeIMO.Email.Store;

internal sealed class PstBlockReference {
    internal PstBlockReference(ulong bid, long offset, int dataLength) {
        Bid = bid;
        Offset = offset;
        DataLength = dataLength;
    }

    internal ulong Bid { get; }
    internal long Offset { get; }
    internal int DataLength { get; }
}

internal sealed class PstNodeReference {
    internal PstNodeReference(uint nid, ulong dataBid, ulong subnodeBid, uint parentNid) {
        Nid = nid;
        DataBid = dataBid;
        SubnodeBid = subnodeBid;
        ParentNid = parentNid;
    }

    internal uint Nid { get; }
    internal ulong DataBid { get; }
    internal ulong SubnodeBid { get; }
    internal uint ParentNid { get; }
    internal int Type => (int)(Nid & 0x1F);
}

internal sealed class PstDataTree {
    internal PstDataTree(IReadOnlyList<byte[]> blocks, long totalLength) {
        Blocks = blocks;
        TotalLength = totalLength;
    }

    internal IReadOnlyList<byte[]> Blocks { get; }
    internal long TotalLength { get; }

    internal byte[] ToArray(long maximum) {
        if (TotalLength > maximum || TotalLength > int.MaxValue) {
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem),
                TotalLength, Math.Min(maximum, int.MaxValue));
        }
        var result = new byte[checked((int)TotalLength)];
        int offset = 0;
        foreach (byte[] block in Blocks) {
            Buffer.BlockCopy(block, 0, result, offset, block.Length);
            offset += block.Length;
        }
        return result;
    }
}

internal sealed class PstSubnodeReference {
    internal PstSubnodeReference(uint nid, ulong dataBid, ulong subnodeBid) {
        Nid = nid;
        DataBid = dataBid;
        SubnodeBid = subnodeBid;
    }

    internal uint Nid { get; }
    internal ulong DataBid { get; }
    internal ulong SubnodeBid { get; }
    internal int Type => (int)(Nid & 0x1F);
}
