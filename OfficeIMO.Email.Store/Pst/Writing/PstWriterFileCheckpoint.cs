namespace OfficeIMO.Email.Store;

internal readonly struct PstWriterFileCheckpoint {
    internal PstWriterFileCheckpoint(long streamLength, long nextOffset,
        ulong nextBlockBid, ulong nextPageBid, long blockCount) {
        StreamLength = streamLength;
        NextOffset = nextOffset;
        NextBlockBid = nextBlockBid;
        NextPageBid = nextPageBid;
        BlockCount = blockCount;
    }
    internal long StreamLength { get; }
    internal long NextOffset { get; }
    internal ulong NextBlockBid { get; }
    internal ulong NextPageBid { get; }
    internal long BlockCount { get; }
}
