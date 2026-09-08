using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderContentTraversalTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void PathContainersStayTogetherBeforeApplyingLocalPositions(bool numbered, bool roundTrip) {
        ReaderLocation Location(string path, int position) => new() { Path = path, SourceBlockIndex = position,
            Page = numbered ? 1 : null, BlockAnchor = "block-" + position };
        var source = new OfficeDocumentReadResult {
            Blocks = new[] { "z.txt", "a.txt" }.SelectMany(path => new[] {
                new OfficeDocumentBlock { Id = path + "1", Text = path + "1", Location = Location(path, 1) },
                new OfficeDocumentBlock { Id = path + "2", Text = path + "2", Location = Location(path, 2) },
                new OfficeDocumentBlock { Id = path + "3", Text = path + "3", Location = Location(path, 3) }
            }).ToArray(),
            Tables = new[] { "z.txt", "a.txt" }.Select(path => new ReaderTable {
                Columns = new[] { "value" }, Rows = new[] { new[] { path } }, Location = Location(path, 2)
            }).ToArray()
        };
        var document = roundTrip ? OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) : source;
        Assert.Equal(new[] { "z.txt1", "z.txt2", "z.txt3", "a.txt1", "a.txt2", "a.txt3" },
            document.EnumerateBlocks().Select(block => block.Text));
        Assert.Equal(new[] { "z.txt1", "z.txt2", "z.txt", "z.txt3", "a.txt1", "a.txt2", "a.txt", "a.txt3" },
            document.EnumerateContent().Select(item => item.Block?.Text ?? item.Table!.Rows[0][0]));
        var hierarchy = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions { MaxTokens = 100, OverlapTokens = 0, IncludeContextInText = false });
        string text = string.Join("\n", hierarchy.Chunks.Select(chunk => chunk.Text));
        Assert.True(text.IndexOf("z.txt3", StringComparison.Ordinal) < text.IndexOf("a.txt1", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("sheet")]
    [InlineData("slide")]
    public void PageResolverKeepsBlocksAndTablesInTheSameNonPageContainer(string kind) {
        var page = new OfficeDocumentPage { Number = 7, Name = "Inventory", Location = new() { SourceBlockKind = kind },
            Blocks = new[] { new OfficeDocumentBlock { Text = "block" } },
            Tables = new[] { new ReaderTable { Columns = new[] { "value" }, Rows = new[] { new[] { "cell" } } } }
        };
        var document = new OfficeDocumentReadResult { Pages = new[] { page } };
        var locations = document.EnumerateContent().Select(item => item.Location!).Append(page.GetResolvedLocation()).ToArray();
        Assert.Equal(3, locations.Length);
        Assert.All(locations, location => {
            Assert.Null(location.Page);
            Assert.Equal(kind == "slide" ? 7 : (int?)null, location.Slide);
            Assert.Equal(kind == "sheet" ? "Inventory" : null, location.Sheet);
        });
        Assert.Null(page.Location.Slide);
        Assert.Null(page.Location.Sheet);
    }

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
                new OfficeDocumentBlock { Text = "placeholder" + container, Location = Location(container, "table-1") },
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
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void AnchorProjectionsUsePageFallbackWithoutLosingOtherContainers(bool ambiguousAggregate, bool roundTrip) {
        var aggregate = new OfficeDocumentBlock { Text = "aggregate", Location = new() {
            Page = ambiguousAggregate ? null : 2, BlockAnchor = "paragraph-1" } };
        var source = new OfficeDocumentReadResult { Blocks = new[] { aggregate }, Pages = new[] {
            new OfficeDocumentPage { Location = new() { Page = 1 }, Blocks = new[] {
                new OfficeDocumentBlock { Text = "page-one", Location = new() { BlockAnchor = "paragraph-1" } } } },
            new OfficeDocumentPage { Location = new() { Page = 2 }, Blocks = new[] {
                new OfficeDocumentBlock { Text = "page-two", Location = new() { BlockAnchor = "paragraph-1" } } } }
        } };
        var document = roundTrip ? OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) : source;
        var blocks = document.EnumerateBlocks().ToArray();
        Assert.Equal(ambiguousAggregate ? new[] { "page-one", "page-two", "aggregate" } : new[] { "page-one", "aggregate" },
            blocks.Select(block => block.Text));
        Assert.Equal(ambiguousAggregate ? (int?)null : 2, aggregate.Location!.Page);
        Assert.All(source.Pages, page => Assert.Null(page.Blocks[0].Location!.Page));
        var hierarchy = ReaderHierarchicalChunker.Chunk(document,
            new ReaderHierarchicalChunkingOptions { MaxTokens = 100, OverlapTokens = 0, IncludeContextInText = false });
        string text = string.Join("\n", hierarchy.Chunks.Select(chunk => chunk.Text));
        Assert.Contains("page-one", text);
        Assert.Contains("aggregate", text);
        if (ambiguousAggregate) Assert.Contains("page-two", text);
        else Assert.DoesNotContain("page-two", text);
        Assert.Equal(ambiguousAggregate ? (int?)null : 2,
            Assert.Single(hierarchy.Chunks, chunk => chunk.Text.Contains("aggregate")).Location.Page);
    }

    [Theory]
    [InlineData("page", false)]
    [InlineData("page", true)]
    [InlineData("slide", false)]
    [InlineData("slide", true)]
    [InlineData("sheet", false)]
    [InlineData("sheet", true)]
    [InlineData("path", false)]
    [InlineData("path", true)]
    public void HierarchyPreservesAuthoritativeAnchorProjectionAtInputLimit(string kind, bool roundTrip) {
        var aggregate = new OfficeDocumentBlock { Text = "authoritative aggregate", Location = new() { BlockAnchor = "paragraph-1" } };
        var source = new OfficeDocumentReadResult { Blocks = new[] { aggregate }, Pages = new[] {
            new OfficeDocumentPage { Number = kind == "path" ? null : 7, Name = "Inventory",
                Location = new() { SourceBlockKind = kind, Path = "source.dat" }, Blocks = new[] {
                new OfficeDocumentBlock { Text = "page copy", Location = new() { BlockAnchor = "paragraph-1", SourceBlockIndex = 12 } }
            } }
        } };
        var document = roundTrip ? OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) : source;
        var result = ReaderHierarchicalChunker.Chunk(document, new ReaderHierarchicalChunkingOptions {
            MaxInputChunks = 1, MaxTokens = 100, OverlapTokens = 0, IncludeContextInText = false
        });
        var chunk = Assert.Single(result.Chunks);
        Assert.Equal("authoritative aggregate", chunk.Text);
        Assert.Equal(kind == "page" ? 7 : (int?)null, chunk.Location.Page);
        Assert.Equal(kind == "slide" ? 7 : (int?)null, chunk.Location.Slide);
        Assert.Equal(kind == "sheet" ? "Inventory" : null, chunk.Location.Sheet);
        Assert.Equal("source.dat", chunk.Location.Path);
        Assert.Equal(12, chunk.Location.SourceBlockIndex);
        Assert.Equal("paragraph-1", chunk.Location.BlockAnchor);
        Assert.Null(aggregate.Location!.Page);
        Assert.Null(aggregate.Location.SourceBlockIndex);
        var canonical = Assert.Single(document.EnumerateBlocks());
        Assert.Equal(chunk.Text, canonical.Text);
        Assert.Equal(chunk.Location.Page, canonical.Location!.Page);
        Assert.Equal(chunk.Location.Slide, canonical.Location.Slide);
        Assert.Equal(chunk.Location.Sheet, canonical.Location.Sheet);
        Assert.Equal(chunk.Location.Path, canonical.Location.Path);
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
