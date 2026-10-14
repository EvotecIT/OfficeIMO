using OfficeIMO.Excel;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed class ExcelUtf8TextCacheTests {
    [Fact]
    public void Utf8Cache_AlternatingCanonicalIndexesRetainEachEncoding() {
        var cache = new ExcelUtf8TextCache();
        string[] names = Enumerable.Range(0, 8).Select(index => "name-" + index).ToArray();
        ArraySegment<byte>[] originals = names.Select((name, index) => cache.Get(index, name)).ToArray();
        for (int row = 0; row < 512; row++) {
            int index = row % names.Length;
            ArraySegment<byte> current = cache.Get(index, names[index]);
            Assert.Same(originals[index].Array, current.Array);
            Assert.Equal(names[index], Decode(current));
        }
    }

    [Fact]
    public void Utf8Cache_ReusedKeyKeepsEachCanonicalStringAndPreviousBorrowedArray() {
        var cache = new ExcelUtf8TextCache();
        ArraySegment<byte> first = cache.Get(0, "first 🐢");
        ArraySegment<byte> second = cache.Get(0, "second 🐢");
        Assert.Equal("first 🐢", Decode(first));
        Assert.Equal("second 🐢", Decode(second));
        Assert.NotSame(first.Array, second.Array);
        Assert.Same(second.Array, cache.Get(0, "second 🐢").Array);
        Parallel.For(0, 1000, index => {
            string value = (index & 1) == 0 ? "first 🐢" : "second 🐢";
            Assert.Equal(value, Decode(cache.Get(0, value)));
        });
        Assert.Equal("first 🐢", Decode(first));
    }

    [Fact]
    public void SharedStringUtf8Cache_InvalidIndexesDoNotResolveAnotherEntry() {
        byte[] payload = Encoding.UTF8.GetBytes(
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><si><t>valid</t></si></sst>");
        var cache = SharedStringCache.Build(() => new MemoryStream(payload, writable: false), new ExcelReadOptions());
        Assert.True(cache.TryGetUtf8(0, out ArraySegment<byte> valid));
        Assert.Equal("valid", Decode(valid));
        foreach (int index in new[] { -1, 1, int.MaxValue }) {
            Assert.False(cache.TryGetUtf8(index, out ArraySegment<byte> invalid));
            Assert.Null(invalid.Array);
            Assert.Equal(0, invalid.Count);
        }
        Assert.True(cache.TryGetUtf8(0, out ArraySegment<byte> repeated));
        Assert.Same(valid.Array, repeated.Array);
    }

    private static string Decode(ArraySegment<byte> text) =>
        Encoding.UTF8.GetString(text.Array!, text.Offset, text.Count);
}
