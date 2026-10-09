using System.Buffers;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed class SharedStringIndexTests {
    [Fact]
    public void IndexedTable_UsesActualCountAndKeepsBorrowedTextUntilDisposal() {
        using OpenXmlPooledPartStream stream = CreateStream(
            "<si><t>first</t></si><si><t></t></si><si><t xml:space=\"preserve\"> </t></si>",
            declaredCount: "2147483647");
        byte[] buffer = stream.BorrowBuffer(out _);
        using SharedStringCache cache = SharedStringCache.Build(() => stream,
            new ExcelReadOptions { MaxSharedStringItems = 3 });

        Assert.Equal(3, cache.Count);
        Assert.True(stream.CanRead);
        Assert.True(cache.TryGetUtf8(0, out ArraySegment<byte> first));
        Assert.Same(buffer, first.Array);
        Assert.Equal("first", Decode(first));
        Assert.True(cache.TryGetUtf8(1, out ArraySegment<byte> empty));
        Assert.Equal(0, empty.Count);
        Assert.True(cache.TryGetUtf8(2, out ArraySegment<byte> space));
        Assert.Equal(" ", Decode(space));
        Assert.Equal("first", cache.Get(0));
        Assert.Equal(string.Empty, cache.Get(1));
        Assert.Equal(" ", cache.Get(2));
        Assert.Equal("first", Decode(first));
        foreach (int index in new[] { -1, 3, int.MaxValue }) {
            Assert.Null(cache.Get(index));
            Assert.False(cache.TryGetUtf8(index, out _));
        }

        cache.Dispose();
        Assert.False(stream.CanRead);
        Assert.Throws<ObjectDisposedException>(() => stream.BorrowBuffer(out _));
        Assert.Throws<ObjectDisposedException>(() => cache.Get(0));
    }

    [Theory]
    [InlineData("Items")]
    [InlineData("ItemCharacters")]
    [InlineData("AggregateCharacters")]
    public void IndexedTable_RejectsUnreferencedTextOverLimitsAndReleasesItsStream(string limit) {
        using OpenXmlPooledPartStream stream = CreateStream(
            "<si><t>first</t></si><si><t>unreferenced</t></si>");
        var options = new ExcelReadOptions {
            MaxSharedStringItems = limit == "Items" ? 1 : 2,
            MaxSharedStringItemCharacters = limit == "ItemCharacters" ? 5 : 20,
            MaxSharedStringCharacters = limit == "AggregateCharacters" ? 5 : 20
        };
        using SharedStringCache cache = SharedStringCache.Build(() => stream, options);
        Assert.Throws<InvalidDataException>(() => cache.EnsureLoaded());
        Assert.False(stream.CanRead);
    }

    [Fact]
    public void IndexedTable_CanceledInitializationReleasesTheOpenedStream() {
        using var cancellation = new CancellationTokenSource();
        using OpenXmlPooledPartStream stream = CreateStream("<si><t>first</t></si>");
        using SharedStringCache cache = SharedStringCache.Build(() => {
            cancellation.Cancel();
            return stream;
        }, new ExcelReadOptions { CancellationToken = cancellation.Token });

        Assert.Throws<OperationCanceledException>(() => cache.EnsureLoaded());
        Assert.False(stream.CanRead);
    }

    [Fact]
    public void IndexedTable_DisposalBeforeInitializationDoesNotOpenThePart() {
        int opens = 0;
        using SharedStringCache cache = SharedStringCache.Build(() => {
            opens++;
            return CreateStream("<si><t>first</t></si>");
        }, new ExcelReadOptions());
        cache.Dispose();
        Assert.Throws<ObjectDisposedException>(() => cache.EnsureLoaded());
        Assert.Equal(0, opens);
    }

    [Fact]
    public void IndexedTable_ConcurrentStringAndUtf8GettersKeepCanonicalValuesIndependent() {
        using OpenXmlPooledPartStream stream = CreateStream(
            "<si><t>first</t></si><si><t>second</t></si>");
        using SharedStringCache cache = SharedStringCache.Build(() => stream, new ExcelReadOptions());
        Assert.True(cache.TryGetUtf8(0, out ArraySegment<byte> first));
        Parallel.For(0, 512, iteration => {
            int index = iteration & 1;
            string expected = index == 0 ? "first" : "second";
            Assert.Equal(expected, cache.Get(index));
            Assert.True(cache.TryGetUtf8(index, out ArraySegment<byte> borrowed));
            Assert.Equal(expected, Decode(borrowed));
        });
        Assert.Equal("first", Decode(first));
        Assert.Equal(new[] { 1 }, cache.FindIndexesContaining("econd", StringComparison.Ordinal));
    }

    private static OpenXmlPooledPartStream CreateStream(string items, string declaredCount = "2") {
        byte[] bytes = Encoding.UTF8.GetBytes(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"" + declaredCount +
            "\" uniqueCount=\"" + declaredCount + "\">" + items + "</sst>");
        byte[] buffer = ArrayPool<byte>.Shared.Rent(bytes.Length);
        Buffer.BlockCopy(bytes, 0, buffer, 0, bytes.Length);
        return new OpenXmlPooledPartStream(buffer, bytes.Length);
    }

    private static string Decode(ArraySegment<byte> text) =>
        Encoding.UTF8.GetString(text.Array!, text.Offset, text.Count);
}
