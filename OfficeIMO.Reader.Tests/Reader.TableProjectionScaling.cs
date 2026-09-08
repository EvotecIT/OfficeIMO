using System.Collections;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderTableProjectionScalingTests {
    [Fact]
    public void DistinctReconciledAliasesInOneChunkStillCountAsTwoOccurrences() {
        ReaderTable Table() => new() { Columns = new[] { "Value" }, Rows = new[] { new[] { "42" } },
            Location = new() { Page = 1, TableIndex = 0 } };
        ReaderTable aggregate = Table(), page = Table();
        foreach (bool aggregateFirst in new[] { false, true }) {
            var document = new OfficeDocumentReadResult {
                Tables = new[] { aggregate },
                Pages = new[] { new OfficeDocumentPage { Number = 1, Tables = new[] { page } } },
                Chunks = new[] { new ReaderChunk { Tables = aggregateFirst ? new[] { aggregate, page } : new[] { page, aggregate } } }
            };
            foreach (var source in new[] { document, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(document)) })
                Assert.Equal(2, source.EnumerateTables().Count());
        }
    }

    [Theory]
    [InlineData("aggregate", false)]
    [InlineData("aggregate", true)]
    [InlineData("page", false)]
    [InlineData("page", true)]
    [InlineData("both", false)]
    [InlineData("both", true)]
    public void SharedChunkReferencesCannotConsumeAnotherEqualOccurrence(string owner, bool sharedFirst) {
        ReaderTable Table() => new() { Columns = new[] { "Value" }, Rows = new[] { new[] { "42" } },
            Location = new() { Page = 1, TableIndex = 0 } };
        ReaderTable shared = Table();
        var document = new OfficeDocumentReadResult {
            Tables = owner == "aggregate" ? new[] { shared } : owner == "both" ? new[] { Table() } : Array.Empty<ReaderTable>(),
            Pages = owner == "aggregate" ? Array.Empty<OfficeDocumentPage>() : new[] { new OfficeDocumentPage { Number = 1, Tables = new[] { shared } } },
            Chunks = new[] { new ReaderChunk { Tables = sharedFirst ? new[] { shared, Table() } : new[] { Table(), shared } } }
        };
        foreach (var source in new[] { document, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(document)) }) {
            Assert.Equal(2, source.EnumerateTables().Count());
            Assert.Equal(2, source.EnumerateContent().Count(item => item.Table != null));
        }
    }

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
