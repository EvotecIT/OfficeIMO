using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderTableProjectionCoverageTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void PartialLocationsCannotConsumeLaterExplicitOccurrences(bool chunks, bool shared) {
        var unscoped = Table(); unscoped.Location = null;
        var first = Table();
        var second = Table(); second.Location!.Page = 2;
        var copy = Table();
        var document = new OfficeDocumentReadResult {
            Tables = chunks ? new[] { first, second } : new[] { unscoped, first },
            Pages = chunks ? System.Array.Empty<OfficeDocumentPage>() : new[] {
                new OfficeDocumentPage { Number = 1, Tables = new[] { shared ? first : copy } },
                new OfficeDocumentPage { Number = 2, Tables = new[] { second } }
            },
            Chunks = chunks ? new[] { new ReaderChunk { Tables = new[] { unscoped, shared ? first : copy } } }
                : System.Array.Empty<ReaderChunk>()
        };
        foreach (var source in new[] { document, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(document)) }) {
            var tables = System.Linq.Enumerable.ToArray(source.EnumerateTables());
            Assert.Equal(2, tables.Length);
            Assert.Equal(new int?[] { 1, 2 }, System.Linq.Enumerable.OrderBy(System.Linq.Enumerable.Select(tables, table => table.Location!.Page), page => page));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IncomparablePartialCoordinatesFindTheRemainingCompatibleOccurrence(bool chunks) {
        var byPage = Table(); byPage.Location = new() { Page = 1 };
        var byIndex = Table(); byIndex.Location = new() { TableIndex = 0 };
        var first = Table();
        var second = Table(); second.Location!.TableIndex = 1;
        var document = new OfficeDocumentReadResult {
            Tables = chunks ? new[] { first, second } : new[] { byPage, byIndex },
            Pages = chunks ? System.Array.Empty<OfficeDocumentPage>() : new[] { new OfficeDocumentPage { Number = 1, Tables = new[] { first, second } } },
            Chunks = chunks ? new[] { new ReaderChunk { Tables = new[] { byPage, byIndex } } } : System.Array.Empty<ReaderChunk>()
        };
        foreach (var source in new[] { document, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(document)) }) {
            var tables = System.Linq.Enumerable.ToArray(source.EnumerateTables());
            Assert.Equal(2, tables.Length);
            Assert.Equal(new int?[] { 0, 1 }, System.Linq.Enumerable.OrderBy(System.Linq.Enumerable.Select(tables, table => table.Location!.TableIndex), index => index));
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void AliasesRetainConservativeCoverageWithoutMutatingSources(bool chunkAlias, bool richerRowCount) {
        var aggregate = Table();
        var alias = Table();
        alias.Truncated = true;
        alias.TotalRowCount = richerRowCount ? 3 : 1;
        alias.Diagnostics = new() { SourceRowCount = 3, ExpectedCellCount = 3, FilledCellCount = 1,
            MissingCellCount = 2, Confidence = 0.4, HasGeometry = true, XStart = 10, XEnd = 50, Width = 40 };
        var source = new OfficeDocumentReadResult {
            Tables = new[] { aggregate },
            Pages = chunkAlias ? System.Array.Empty<OfficeDocumentPage>() : new[] { new OfficeDocumentPage { Number = 1, Tables = new[] { alias } } },
            Chunks = chunkAlias ? new[] { new ReaderChunk { Location = new() { Page = 1 }, Tables = new[] { alias } } } : System.Array.Empty<ReaderChunk>()
        };
        foreach (var document in new[] { source, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) }) {
            var table = Assert.Single(document.EnumerateTables());
            Assert.True(table.Truncated);
            Assert.Equal(richerRowCount ? 3 : 1, table.TotalRowCount);
            Assert.NotNull(table.Diagnostics);
            Assert.Equal(2, table.Diagnostics!.MissingCellCount);
            Assert.Equal(3, table.Diagnostics.SourceRowCount);
            Assert.True(table.Diagnostics.HasGeometry);
            Assert.Equal(10, table.Diagnostics.XStart);
        }
        Assert.False(aggregate.Truncated);
        Assert.Null(aggregate.Diagnostics);
        Assert.Equal(1, aggregate.TotalRowCount);
        Assert.True(alias.Truncated);
    }

    private static ReaderTable Table() => new() {
        Columns = new[] { "Item" }, Rows = new[] { new[] { "Widget" } }, TotalRowCount = 1,
        Location = new() { Page = 1, TableIndex = 0 }
    };

    [Fact]
    public void ConflictingAliasDiagnosticsRetainTheLowerCoverageEstimate() {
        var aggregate = Table();
        aggregate.Diagnostics = new() { SourceRowCount = 1, ExpectedCellCount = 1, FilledCellCount = 1,
            Confidence = 1, SchemaConfidence = 1, CellCompleteness = 1, ColumnGeometryConfidence = 1 };
        var alias = Table();
        alias.Diagnostics = new() { SourceRowCount = 3, ExpectedCellCount = 3, FilledCellCount = 1,
            MissingCellCount = 2, Confidence = 0.4, SchemaConfidence = 0.5, CellCompleteness = 0.3,
            HasGeometry = true, XStart = 10, XEnd = 50, Width = 40 };
        var source = new OfficeDocumentReadResult { Tables = new[] { aggregate }, Pages = new[] {
            new OfficeDocumentPage { Number = 1, Tables = new[] { alias } }
        } };
        foreach (var document in new[] { source, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) }) {
            var diagnostics = Assert.Single(document.EnumerateTables()).Diagnostics!;
            Assert.Equal(3, diagnostics.SourceRowCount);
            Assert.Equal(2, diagnostics.MissingCellCount);
            Assert.Equal(0.4, diagnostics.Confidence);
            Assert.Equal(0.5, diagnostics.SchemaConfidence);
            Assert.Equal(0.3, diagnostics.CellCompleteness);
            Assert.True(diagnostics.HasGeometry);
            Assert.Equal(10, diagnostics.XStart);
        }
        Assert.Equal(1, aggregate.Diagnostics.Confidence);
        Assert.False(aggregate.Diagnostics.HasGeometry);
        Assert.Equal(0.4, alias.Diagnostics.Confidence);
    }
}
