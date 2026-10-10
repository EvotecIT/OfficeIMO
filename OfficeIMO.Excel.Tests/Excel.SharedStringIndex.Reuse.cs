using System.Globalization;
using System.Threading;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed partial class SharedStringIndexTests {
    [Fact]
    public void IndexedTable_CompletedOwnersAndOverlappingReadersKeepExactIndependentValues() {
        string[] seedValues = CreateReuseIndexedValues("seed-", 12_291);
        using OpenXmlPooledPartStream seedStream = CreateStream(seedValues);
        using SharedStringCache seed = SharedStringCache.Build(() => seedStream, new ExcelReadOptions());
        AssertExactIndexedValues(seed, seedValues);
        seed.Dispose();
        AssertIndexedOwnerClosed(seed, seedStream);

        string[] firstValues = CreateReuseIndexedValues("A-", 8_195);
        using OpenXmlPooledPartStream firstStream = CreateStream(firstValues);
        using SharedStringCache first = SharedStringCache.Build(() => firstStream, new ExcelReadOptions());
        Assert.Equal(firstValues.Length, first.Count);
        Assert.True(first.TryGetUtf8(firstValues.Length - 1, out ArraySegment<byte> borrowedFirst));
        byte[] expectedBorrowedFirst = Encoding.UTF8.GetBytes(firstValues[firstValues.Length - 1]);
        // A completed old owner cannot return storage currently used by the next reader.
        seed.Dispose();

        string[] secondValues = CreateReuseIndexedValues("second-longer-", 9_219);
        using OpenXmlPooledPartStream secondStream = CreateStream(secondValues);
        using SharedStringCache second = SharedStringCache.Build(() => secondStream, new ExcelReadOptions());
        AssertExactIndexedValues(second, secondValues);
        AssertExactIndexedValues(first, firstValues);
        Assert.Equal(expectedBorrowedFirst, borrowedFirst.ToArray());
        second.Dispose();
        second.Dispose();
        AssertIndexedOwnerClosed(second, secondStream);

        string[] thirdValues = CreateReuseIndexedValues("C-more-text-", 8_197);
        using OpenXmlPooledPartStream thirdStream = CreateStream(thirdValues);
        using SharedStringCache third = SharedStringCache.Build(() => thirdStream, new ExcelReadOptions());
        AssertExactIndexedValues(third, thirdValues);
        // Reusing a different completed owner must not mutate the still-live first cache.
        second.Dispose();
        AssertExactIndexedValues(first, firstValues);
        Assert.Equal(expectedBorrowedFirst, borrowedFirst.ToArray());
        third.Dispose();
        AssertIndexedOwnerClosed(third, thirdStream);
        first.Dispose();
        AssertIndexedOwnerClosed(first, firstStream);
    }

    [Theory]
    [InlineData("LateRichText")]
    [InlineData("Items")]
    [InlineData("Cancellation")]
    public void IndexedTable_LargeInitializationExitPreservesRecoveryAndReleasesOpenedStorage(string exit) {
        string[] seedValues = CreateReuseIndexedValues("seed-long-", 12_291);
        using OpenXmlPooledPartStream seedStream = CreateStream(seedValues);
        using SharedStringCache seed = SharedStringCache.Build(() => seedStream, new ExcelReadOptions());
        Assert.Equal(seedValues.Length, seed.Count);
        seed.Dispose();
        AssertIndexedOwnerClosed(seed, seedStream);

        const int failingCount = 8_213;
        string[] expected = CreateReuseIndexedValues("prefix-", failingCount);
        var items = new StringBuilder();
        for (int index = 0; index < expected.Length - 1; index++) {
            items.Append("<si><t xml:space=\"preserve\">").Append(expected[index]).Append("</t></si>");
        }
        if (exit == "LateRichText") {
            expected[expected.Length - 1] = "late rich text";
            items.Append("<si><r><t>late </t></r><r><t>rich text</t></r></si>");
        } else {
            items.Append("<si><t xml:space=\"preserve\">").Append(expected[expected.Length - 1]).Append("</t></si>");
        }
        using OpenXmlPooledPartStream failedStream = CreateStream(items.ToString(), failingCount.ToString(CultureInfo.InvariantCulture));
        using var cancellation = new CancellationTokenSource();
        var options = new ExcelReadOptions {
            MaxSharedStringItems = exit == "Items" ? failingCount - 1 : failingCount,
            CancellationToken = cancellation.Token
        };
        using SharedStringCache failed = SharedStringCache.Build(() => {
            // The actual stream factory boundary gives deterministic cancellation
            // after opening, without clocks, product hooks, or a racing task.
            if (exit == "Cancellation") cancellation.Cancel();
            return failedStream;
        }, options);
        if (exit == "LateRichText") {
            AssertExactIndexedValues(failed, expected);
        } else if (exit == "Items") {
            Assert.Throws<InvalidDataException>(() => failed.EnsureLoaded());
        } else {
            Assert.Throws<OperationCanceledException>(() => failed.EnsureLoaded());
        }
        // The canonical XML fallback owns strings, not the rejected byte stream.
        Assert.False(failedStream.CanRead);
        Assert.Throws<ObjectDisposedException>(() => failedStream.BorrowBuffer(out _));
        failed.Dispose();
        failed.Dispose();

        string[] healthyValues = CreateReuseIndexedValues("recovered-different-", 8_201);
        using OpenXmlPooledPartStream healthyStream = CreateStream(healthyValues);
        using SharedStringCache healthy = SharedStringCache.Build(() => healthyStream, new ExcelReadOptions());
        seed.Dispose();
        failed.Dispose();
        AssertExactIndexedValues(healthy, healthyValues);
        healthy.Dispose();
        AssertIndexedOwnerClosed(healthy, healthyStream);
    }

    private static string[] CreateReuseIndexedValues(string prefix, int count) {
        string[] values = Enumerable.Range(0, count)
            .Select(index => prefix + index.ToString(CultureInfo.InvariantCulture)).ToArray();
        values[count - 2] = string.Empty;
        values[count - 1] = " " + prefix + "preserved ";
        return values;
    }

    private static void AssertExactIndexedValues(SharedStringCache cache, string[] expected) {
        Assert.Equal(expected.Length, cache.Count);
        for (int index = expected.Length - 1; index >= 0; index--) {
            Assert.Equal(expected[index], cache.Get(index));
            Assert.True(cache.TryGetUtf8(index, out ArraySegment<byte> text));
            Assert.Equal(Encoding.UTF8.GetBytes(expected[index]), text.ToArray());
        }
        foreach (int index in new[] { -1, expected.Length, int.MaxValue }) {
            Assert.Null(cache.Get(index));
            Assert.False(cache.TryGetUtf8(index, out _));
        }
    }

    private static void AssertIndexedOwnerClosed(SharedStringCache cache, OpenXmlPooledPartStream stream) {
        Assert.False(stream.CanRead);
        Assert.Throws<ObjectDisposedException>(() => stream.BorrowBuffer(out _));
        Assert.Throws<ObjectDisposedException>(() => cache.Get(0));
    }
}
