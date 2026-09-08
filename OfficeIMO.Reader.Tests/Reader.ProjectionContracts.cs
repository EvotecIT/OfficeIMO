using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderProjectionContractTests {
    [Theory]
    [InlineData("SourceBlockIndex")]
    [InlineData("BlockIndex")]
    [InlineData("StartLine")]
    [InlineData("Path")]
    [InlineData("BlockAnchor")]
    [InlineData("HeadingPath")]
    [InlineData("HierarchyHeadingPath")]
    public void ChunkCoordinatesFillMissingProjectionInformationButDoNotOverrideConflicts(string field) {
        var property = typeof(ReaderLocation).GetProperty(field)!;
        foreach (bool conflicting in new[] { false, true }) {
            var location = new ReaderLocation { Page = 1 };
            if (conflicting) property.SetValue(location, property.PropertyType == typeof(string) ? "first" : 4);
            var pageTable = new ReaderTable { Location = location, Columns = new[] { "Count" }, Rows = new[] { new[] { "42" } } };
            var chunkLocation = new ReaderLocation { Page = 1 };
            property.SetValue(chunkLocation, property.PropertyType == typeof(string) ? "second" : 5);
            var chunkTable = conflicting ? new ReaderTable { Location = null, Columns = pageTable.Columns, Rows = pageTable.Rows } : pageTable;
            var source = new OfficeDocumentReadResult {
                Pages = new[] { new OfficeDocumentPage { Number = 1, Tables = new[] { pageTable } } },
                Chunks = new[] { new ReaderChunk { Location = chunkLocation, Tables = new[] { chunkTable } } }
            };
            foreach (bool roundTrip in new[] { false, true }) {
                var document = roundTrip ? OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) : source;
                Assert.Equal(conflicting ? 2 : 1, document.EnumerateTables().Count());
            }
        }
    }

    [Theory]
    [InlineData("sheet", false)]
    [InlineData("sheet", true)]
    [InlineData("named-sheet", false)]
    [InlineData("named-sheet", true)]
    [InlineData("path", false)]
    [InlineData("path", true)]
    public void AllBlockConsumersUseDeclaredContainerOrder(string kind, bool roundTrip) {
        ReaderLocation Location(string value) => new() { Path = kind == "path" ? value : null, Sheet = kind != "path" ? value : null };
        OfficeDocumentPage Page(string value) => new() { Name = value, Location = kind == "named-sheet" ? new() { SourceBlockKind = "sheet" } : Location(value) };
        var source = new OfficeDocumentReadResult { Pages = new[] { Page("Z"), Page("A") }, Blocks = new[] {
            new OfficeDocumentBlock { Id = "A", Text = "A", Location = Location("A") },
            new OfficeDocumentBlock { Id = "Z", Text = "Z", Location = Location("Z") }
        } };
        var document = roundTrip ? OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) : source;
        Assert.Equal(new[] { "Z", "A" }, document.EnumerateContent().Select(item => item.Block!.Text));
        Assert.Equal(new[] { "Z", "A" }, document.EnumerateBlocks().Select(block => block.Text));
        var hierarchy = ReaderHierarchicalChunker.Chunk(document, new ReaderHierarchicalChunkingOptions {
            MaxInputChunks = 1, MaxTokens = 100, OverlapTokens = 0, IncludeContextInText = false
        });
        Assert.Equal("Z", Assert.Single(hierarchy.Chunks).Text);
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
    public void PageAndChunkTableCopiesKeepMultiplicityAfterTransport(string kind, bool aggregate) {
        ReaderLocation Location(int value) => new() { Page = kind == "page" ? value : null,
            Slide = kind == "slide" ? value : null, Sheet = kind == "sheet" ? "Sheet" + value : null,
            Path = kind == "path" ? "file" + value : null };
        var tables = Enumerable.Range(0, 4).Select(_ => new ReaderTable { Location = null,
            Columns = new[] { "Count" }, Rows = new[] { new[] { "42" } } }).ToArray();
        var source = new OfficeDocumentReadResult {
            Tables = aggregate ? tables : Array.Empty<ReaderTable>(),
            Pages = Enumerable.Range(1, 2).Select(value => new OfficeDocumentPage { Location = Location(value),
                Tables = tables.Skip((value - 1) * 2).Take(2).ToArray() }).ToArray(),
            Chunks = Enumerable.Range(1, 2).Select(value => new ReaderChunk { Location = Location(value),
                Tables = tables.Skip((value - 1) * 2).Take(2).ToArray() }).ToArray()
        };
        foreach (bool roundTrip in new[] { false, true }) {
            var document = roundTrip ? OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) : source;
            Assert.Equal(4, document.EnumerateTables().Count());
            Assert.Equal(4, document.EnumerateContent().Count(item => item.Table != null));
        }
        Assert.All(tables, table => Assert.Null(table.Location));
    }
}
