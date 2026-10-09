using System.Buffers;
using System.Globalization;
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

    [Fact]
    public void IndexedTable_MaterializesSparseAndNonSequentialValuesWithoutChangingBorrowedText() {
        string[] expected = CreateIndexedValues(4099);
        using OpenXmlPooledPartStream stream = CreateStream(expected, declaredCount: "1");
        using SharedStringCache cache = SharedStringCache.Build(() => stream, new ExcelReadOptions());
        Assert.Equal(expected.Length, cache.Count);
        Assert.True(cache.TryGetUtf8(expected.Length - 1, out ArraySegment<byte> last));

        foreach (int index in new[] { 4098, 2048, 17, 1024, 3072, 0, 4097, 2048, 0, 4098 }) {
            Assert.Equal(expected[index], cache.Get(index));
        }
        for (int index = expected.Length - 1; index >= 0; index--) {
            Assert.Equal(expected[index], cache.Get(index));
        }

        Assert.Null(cache.Get(expected.Length));
        Assert.Equal(expected[expected.Length - 1], Decode(last));
        Assert.Equal(expected.Length, cache.Count);
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
        string[] expected = CreateIndexedValues(4099);
        using OpenXmlPooledPartStream stream = CreateStream(expected);
        using SharedStringCache cache = SharedStringCache.Build(() => stream, new ExcelReadOptions());
        Assert.True(cache.TryGetUtf8(0, out ArraySegment<byte> first));
        Parallel.For(0, expected.Length * 3, iteration => {
            int index = expected.Length - 1 - iteration % expected.Length;
            Assert.Equal(expected[index], cache.Get(index));
            Assert.True(cache.TryGetUtf8(index, out ArraySegment<byte> borrowed));
            Assert.Equal(expected[index], Decode(borrowed));
        });
        Assert.Equal(expected[0], Decode(first));
        Assert.Equal(new[] { 2048 }, cache.FindIndexesContaining(expected[2048], StringComparison.Ordinal));
    }

    private static string[] CreateIndexedValues(int count) {
        string[] values = Enumerable.Range(0, count)
            .Select(index => "item-" + index.ToString(CultureInfo.InvariantCulture)).ToArray();
        values[count - 2] = string.Empty;
        values[count - 1] = " preserved ";
        return values;
    }

    private static OpenXmlPooledPartStream CreateStream(string[] values, string? declaredCount = null) {
        var items = new StringBuilder();
        foreach (string value in values) {
            items.Append("<si><t xml:space=\"preserve\">").Append(value).Append("</t></si>");
        }
        return CreateStream(items.ToString(), declaredCount ?? values.Length.ToString(CultureInfo.InvariantCulture));
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
