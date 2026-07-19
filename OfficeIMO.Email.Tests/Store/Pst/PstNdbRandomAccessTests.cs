namespace OfficeIMO.Email.Store.Tests;

public sealed class PstNdbRandomAccessTests {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, true)]
    public void Resolves_nodes_and_blocks_without_loading_complete_indexes(
        bool ansi, bool fourK, bool compressed) {
        using var stream = new MemoryStream(PstTestFileBuilder.Create(
            ost: fourK, ansi: ansi, fourK: fourK, compressBlocks: compressed));
        EmailStoreFormat format = fourK ? EmailStoreFormat.Ost : EmailStoreFormat.Pst;
        PstHeader header = PstHeader.Read(stream, format);
        var reader = new PstNdbReader(stream, header, EmailStoreReaderOptions.Default,
            System.Threading.CancellationToken.None);

        Assert.Throws<InvalidOperationException>(() => reader.Nodes.Count);
        Assert.Equal(4, reader.EnumerateNodes().Count());
        Assert.True(reader.TryGetNode(0x8004, out PstNodeReference? item));
        Assert.NotNull(item);

        PstDataTree data = reader.ReadDataTree(item!.DataBid, 1024 * 1024);

        Assert.NotEmpty(data.Blocks);
        Assert.True(data.TotalLength > 0);
        Assert.False(reader.TryGetNode(0xDEADBEE0, out _));
    }

    [Fact]
    public void Page_cache_is_bounded_and_uses_recent_entries() {
        var cache = new PstPageCache(2);
        int reads = 0;

        cache.GetOrAdd(1, () => CreatePage(ref reads, 1));
        cache.GetOrAdd(2, () => CreatePage(ref reads, 2));
        cache.GetOrAdd(1, () => CreatePage(ref reads, 1));
        cache.GetOrAdd(3, () => CreatePage(ref reads, 3));
        cache.GetOrAdd(2, () => CreatePage(ref reads, 2));

        Assert.Equal(4, reads);
    }

    private static byte[] CreatePage(ref int reads, byte value) {
        reads++;
        return new[] { value };
    }
}
