using System.Collections;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderTableProjectionScalingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void RepeatedTablesReadPayloadsLinearlyAndRetainLocations(bool chunks, bool conflicts) {
        const int count = 400;
        var cells = new CountedCells();
        ReaderTable Table(int index) => new() { Columns = new[] { "Value" }, Rows = new[] { cells },
            Location = new() { Path = "document", Page = 1, SourceBlockIndex = index, TableIndex = index } };
        var document = new OfficeDocumentReadResult { Tables = Enumerable.Range(0, count).Select(Table).ToArray() };
        var copies = Enumerable.Range(conflicts ? count : 0, count).Reverse().Select(Table).ToArray();
        if (chunks) document.Chunks = new[] { new ReaderChunk { Tables = copies } };
        else document.Pages = new[] { new OfficeDocumentPage { Number = 1, Tables = copies } };

        ReaderTable[] result = document.EnumerateTables().ToArray();

        Assert.Equal(conflicts ? count * 2 : count, result.Length);
        Assert.Equal(Enumerable.Range(0, count), result.Take(count).Select(table => table.Location!.SourceBlockIndex!.Value));
        Assert.All(result, table => Assert.Equal(1, table.Location!.Page));
        // Count actual caller-owned cell accesses, independent of elapsed time or private index details.
        Assert.InRange(cells.Reads, 1, count * 8);
        foreach (ReaderTable table in result) Assert.Same(cells, Assert.Single(table.Rows));
    }

    private sealed class CountedCells : IReadOnlyList<string> {
        internal int Reads;
        public int Count => 1;
        public string this[int index] { get { Reads++; return index == 0 ? "42" : throw new ArgumentOutOfRangeException(nameof(index)); } }
        public IEnumerator<string> GetEnumerator() { yield return this[0]; }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
