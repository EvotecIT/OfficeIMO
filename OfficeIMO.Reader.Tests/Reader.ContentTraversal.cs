using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderContentTraversalTests {
    [Theory]
    [InlineData("page", false)]
    [InlineData("page", true)]
    [InlineData("slide", false)]
    [InlineData("slide", true)]
    [InlineData("sheet", false)]
    [InlineData("sheet", true)]
    [InlineData("path", false)]
    [InlineData("path", true)]
    public void ReusedAnchorsRetainTablePositionsWithinTheirOwnContainers(string kind, bool roundTrip) {
        ReaderLocation Location(int container, string? anchor = null) => new() {
            Page = kind == "page" ? container : null, Slide = kind == "slide" ? container : null,
            Sheet = kind == "sheet" ? "Sheet" + container : null, Path = kind == "path" ? "source" + container : null,
            BlockAnchor = anchor
        };
        var source = new OfficeDocumentReadResult {
            Blocks = Enumerable.Range(1, 2).SelectMany(container => new[] {
                new OfficeDocumentBlock { Id = "before" + container, Text = "before" + container, Location = Location(container) },
                new OfficeDocumentBlock { Id = "placeholder" + container, Text = "placeholder" + container, Location = Location(container, "table-1") },
                new OfficeDocumentBlock { Id = "after" + container, Text = "after" + container, Location = Location(container) }
            }).ToArray(),
            Tables = Enumerable.Range(1, 2).Select(container => new ReaderTable {
                Columns = new[] { "Count" }, Rows = new[] { new[] { container.ToString() } }, Location = Location(container, "table-1")
            }).ToArray()
        };
        var document = roundTrip ? OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) : source;
        Assert.Equal(new[] { "before1", "placeholder1", "table", "after1", "before2", "placeholder2", "table", "after2" },
            document.EnumerateContent().Select(item => item.Block?.Text ?? "table"));
        Assert.All(source.Tables, table => Assert.Null(table.Location!.SourceBlockIndex));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StableAnchorsKeepTablesWithTheirPlaceholdersWithoutInventingPositions(bool roundTrip) {
        var source = new OfficeDocumentReadResult {
            Blocks = new[] { Block("before"), Block("placeholder", "table-anchor"), Block("after") },
            Tables = new[] { Table("table-anchor") }
        };
        var document = roundTrip ? OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) : source;
        var content = document.EnumerateContent().ToArray();
        Assert.Equal(new[] { "before", "placeholder", "table", "after" }, content.Select(item => item.Block?.Text ?? "table"));
        Assert.Null(content[2].Location!.SourceBlockIndex);
        Assert.Null(content[2].Location!.BlockIndex);
        Assert.Single(content[2].Table!.Rows);
        Assert.Null(source.Tables[0].Location!.SourceBlockIndex);
    }

    [Fact]
    public void MixedSheetContentUsesTheSourceContainerOrder() {
        var z = Block("Z paragraph"); z.Location = new() { Sheet = "Z", SourceBlockIndex = 3 };
        var a = Block("A paragraph"); a.Location = new() { Sheet = "A", SourceBlockIndex = 3 };
        var zTable = Table(); zTable.Location = new() { Sheet = "Z", SourceBlockIndex = 2 };
        var aTable = Table(); aTable.Location = new() { Sheet = "A", SourceBlockIndex = 2 };
        var document = new OfficeDocumentReadResult { Blocks = new[] { a, z }, Tables = new[] { aTable, zTable }, Pages = new[] {
            new OfficeDocumentPage { Location = new() { Sheet = "Z" } },
            new OfficeDocumentPage { Location = new() { Sheet = "A" } }
        } };
        var content = document.EnumerateContent().ToArray();
        Assert.Equal(new[] { "Z", "Z", "A", "A" }, content.Select(item => item.Location!.Sheet));
        Assert.NotNull(content[0].Table);
        Assert.NotNull(content[1].Block);
        Assert.NotNull(content[2].Table);
        Assert.NotNull(content[3].Block);
    }

    [Fact]
    public void EmptyBlocksRetainChunkTextAndChunkOnlyTableLocation() {
        var table = Table(); table.Location = null;
        var document = new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Location = new() { Page = 2 } } },
            Chunks = new[] { new ReaderChunk { Text = "Chunk text", Location = new() { Page = 1 }, Tables = new[] { table } } } };
        var content = document.EnumerateContent().ToArray();
        Assert.Equal("Chunk text", content[0].Chunk!.Text);
        Assert.Same(table, content[1].Table);
        Assert.Equal(1, content[1].Location!.Page);
        Assert.NotNull(content[2].Block);
        Assert.Null(table.Location);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AmbiguousOrDifferentPageAnchorsDoNotSupplyAnUnprovenPosition(bool duplicate) {
        var block = Block("first", "anchor"); block.Location.SourceBlockIndex = 1;
        var second = Block("second", "anchor"); second.Location.SourceBlockIndex = 2;
        var table = Table("anchor"); table.Location!.Page = duplicate ? 1 : 2;
        var document = new OfficeDocumentReadResult { Blocks = duplicate ? new[] { block, second } : new[] { block }, Tables = new[] { table } };
        var item = Assert.Single(document.EnumerateContent(), item => item.Table is not null);
        Assert.Null(item.Location!.SourceBlockIndex);
        Assert.Equal(duplicate ? 1 : 2, item.Location.Page);
    }

    private static OfficeDocumentBlock Block(string text, string? anchor = null) => new() {
        Id = text, Text = text, Location = new() { Page = 1, BlockAnchor = anchor }
    };
    private static ReaderTable Table(string? anchor = null) => new() {
        Columns = new[] { "Count" }, Rows = new[] { new[] { "12" } }, Location = new() { Page = 1, BlockAnchor = anchor }
    };
}
