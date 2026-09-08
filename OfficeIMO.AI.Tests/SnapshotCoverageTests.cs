using System.Text.Json;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class SnapshotCoverageTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EqualTablesRetainOccurrenceCountsAndPageOnlyContent(bool roundTrip) {
        ReaderTable Table(int? page = null) => new() { Columns = new[] { "Item" }, Rows = new[] { new[] { "Widget" } },
            Location = page.HasValue ? new ReaderLocation { Page = page } : null };
        var first = Table();
        var second = Table();
        var source = new OfficeDocumentReadResult { Tables = new[] { first, second }, Pages = new[] {
            new OfficeDocumentPage { Number = 3, Tables = new[] { roundTrip ? Table(3) : first } },
            new OfficeDocumentPage { Number = 5, Tables = new[] { roundTrip ? Table(5) : second } },
            new OfficeDocumentPage { Number = 7, Tables = new[] { Table() } }
        } };
        var restored = roundTrip ? OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) : source;
        var snapshot = OfficeAiDocument.FromReadResult(new byte[] { 1 }, restored);
        Assert.Equal(new int?[] { 3, 3, 5, 5, 7, 7 }, snapshot.Evidence.Select(item => item.Page));
        Assert.Null(first.Location);
        Assert.Null(second.Location);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LocationlessTableProjectionsKeepOnePagedObservationAfterTransport(bool roundTrip) {
        var table = new ReaderTable { Columns = new[] { "Item" }, Rows = new[] { new[] { "Widget" } } };
        var source = new OfficeDocumentReadResult { Tables = new[] { table },
            Pages = new[] { new OfficeDocumentPage { Number = 3, Tables = new[] { table } } } };
        var restored = roundTrip ? OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) : source;
        var snapshot = OfficeAiDocument.FromReadResult(new byte[] { 1 }, restored);
        Assert.Equal(2, snapshot.Evidence.Count);
        Assert.All(snapshot.Evidence, item => Assert.Equal(3, item.Page));
        Assert.Null(table.Location);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TransportedBlockIdentityRetainsPageFallbackAndSelection(bool useAnchor) {
        var first = new OfficeDocumentBlock { Id = useAnchor ? "" : "first", Text = "First page",
            Location = new() { BlockAnchor = useAnchor ? "first-anchor" : null } };
        var second = new OfficeDocumentBlock { Id = "second", Text = "Second page", Location = new() { Page = 2 } };
        var source = new OfficeDocumentReadResult { Blocks = new[] { second, first }, Pages = new[] {
            new OfficeDocumentPage { Number = 1, Blocks = new[] { first } }
        } };
        var restored = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source));
        var captured = OfficeAiDocument.FromReadResult(new byte[] { 1 }, restored);
        Assert.Equal(new[] { "First page", "Second page" }, captured.Evidence.Select(item => item.Text));
        Assert.Equal(new int?[] { 1, 2 }, captured.Evidence.Select(item => item.Page));
        var executor = new Capture();
        await new OfficeAiEngine(executor).RunAsync(captured, new() { Instruction = "Read", Pages = new[] { 1 } });
        string input = Assert.Single(executor.Requests).InputJson;
        Assert.Contains("First page", input);
        Assert.DoesNotContain("Second page", input);
        Assert.Null(first.Location.Page);
        Assert.Null(restored.Blocks.Single(item => item.Text == "First page").Location.Page);
    }

    [Fact]
    public void AnonymousRepeatedBlocksRemainSeparateWhileStableAnchorsAreDeduplicated() {
        var first = new OfficeDocumentBlock { Text = "Repeat", Location = new() { Page = 1 } };
        var second = new OfficeDocumentBlock { Text = "Repeat", Location = new() { Page = 1 } };
        var anchored = new OfficeDocumentBlock { Text = "Anchored", Location = new() { Page = 1, BlockAnchor = "a" } };
        var anchoredCopy = new OfficeDocumentBlock { Text = "Anchored", Location = new() { Page = 1, BlockAnchor = "a" } };
        var source = new OfficeDocumentReadResult { Blocks = new[] { first, second, anchored },
            Pages = new[] { new OfficeDocumentPage { Number = 1, Blocks = new[] { first, anchoredCopy } } } };
        Assert.Equal(new[] { "Repeat", "Repeat", "Anchored" }, source.EnumerateBlocks().Select(block => block.Text));
        Assert.Equal(3, OfficeAiDocument.FromReadResult(new byte[] { 1 }, source).Evidence.Count);
    }

    [Fact]
    public void JsonRoundTripKeepsOneObservationPerLogicalBlockAndRetainsPageOnlyContent() {
        var shared = new OfficeDocumentBlock { Id = "shared", Text = "Shared", Location = new() { Page = 1 } };
        var pageOnly = new OfficeDocumentBlock { Id = "page-only", Text = "Page only", Location = new() { Page = 1 } };
        var source = new OfficeDocumentReadResult { Blocks = new[] { shared }, Pages = new[] {
            new OfficeDocumentPage { Number = 1, Blocks = new[] { shared, pageOnly } }
        } };
        var restored = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source));
        var captured = OfficeAiDocument.FromReadResult(new byte[] { 1 }, restored, limits: new() { MaxDocumentBlocks = 2 });
        Assert.Equal(new[] { "shared", "page-only" }, captured.Evidence.Select(item => item.SourceBlockId));
        Assert.Equal(2, restored.EnumerateBlocks().Count());
    }

    [Fact]
    public void WorksheetsKeepSourceContainerOrderAndSortPositionsWithinEachSheet() {
        OfficeDocumentBlock Block(string sheet, int index) => new() { Id = sheet + index, Text = sheet + index,
            Location = new() { Sheet = sheet, SourceBlockIndex = index } };
        var z2 = Block("Z", 2); var z1 = Block("Z", 1); var a1 = Block("A", 1);
        var source = new OfficeDocumentReadResult { Blocks = new[] { z2, a1, z1 }, Pages = new[] {
            new OfficeDocumentPage { Name = "Z", Location = new() { Sheet = "Z" }, Blocks = new[] { z2, z1 } },
            new OfficeDocumentPage { Name = "A", Location = new() { Sheet = "A" }, Blocks = new[] { a1 } }
        } };
        Assert.Equal(new[] { "Z1", "Z2", "A1" }, source.EnumerateBlocks().Select(block => block.Id));
        Assert.Equal(new[] { "Z1", "Z2", "A1" }, OfficeAiDocument.FromReadResult(new byte[] { 1 }, source).Evidence.Select(item => item.Text));
    }

    [Fact]
    public async Task UntitledTableRowsRetainTheirOriginalSectionPlaceholderAnchors() {
        OfficeDocumentBlock Block(string id, string text, string? anchor = null) => new() {
            Id = id, Kind = anchor is null ? "heading" : "table", Text = text, Location = new() { Page = 1, BlockAnchor = anchor }
        };
        ReaderTable Table(string anchor) => new() { Location = new() { Page = 1, BlockAnchor = anchor },
            Columns = new[] { "Count" }, Rows = new[] { new[] { "12" } } };
        var source = new OfficeDocumentReadResult {
            Blocks = new[] { Block("north", "North"), Block("north-table", "Detected table", "north-anchor"),
                Block("south", "South"), Block("south-table", "Detected table", "south-anchor") },
            Tables = new[] { Table("north-anchor"), Table("south-anchor") }
        };
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, source);
        var executor = new Capture();
        await new OfficeAiEngine(executor).RunAsync(document, new() { Operation = OfficeAiOperation.Parse, Instruction = "Read each regional table" });
        using var input = JsonDocument.Parse(Assert.Single(executor.Requests).InputJson);
        var evidence = input.RootElement.GetProperty("evidence").EnumerateArray().ToArray();
        foreach (string region in new[] { "north", "south" }) {
            var placeholder = Assert.Single(evidence, item => item.GetProperty("sourceBlockId").GetString() == region + "-table");
            var row = Assert.Single(evidence, item => item.GetProperty("kind").GetString() == "table-row"
                && item.GetProperty("sourceAnchor").GetString() == region + "-anchor");
            Assert.Equal(placeholder.GetProperty("sourceAnchor").GetString(), row.GetProperty("sourceAnchor").GetString());
        }
        source.Tables[0].Location!.BlockAnchor = "changed";
        Assert.Contains(document.Evidence, item => item.Kind == "table-row" && item.SourceAnchor == "north-anchor");
        Assert.NotEqual(document.SnapshotHash, OfficeAiDocument.FromReadResult(new byte[] { 1 }, source).SnapshotHash);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageScopedBlocksSortBeforeLaterAggregatePagesWithoutMutatingReaderObjects(bool nullLocation) {
        var first = new OfficeDocumentBlock { Id = "first", Text = "FIRST" };
        if (nullLocation) first.Location = null!;
        var originalLocation = first.Location;
        var second = new OfficeDocumentBlock { Id = "second", Text = "SECOND", Location = new() { Page = 2 } };
        var source = new OfficeDocumentReadResult { Blocks = new[] { second }, Pages = new[] {
            new OfficeDocumentPage { Number = 1, Blocks = new[] { first } },
            new OfficeDocumentPage { Number = 2, Blocks = new[] { second } }
        } };
        Assert.Equal(new[] { "first", "second" }, source.EnumerateBlocks().Select(block => block.Id));
        var captured = OfficeAiDocument.FromReadResult(new byte[] { 1 }, source);
        Assert.Equal(new[] { "first", "second" }, captured.Evidence.Select(item => item.SourceBlockId));
        Assert.Equal(new int?[] { 1, 2 }, captured.Evidence.Select(item => item.Page));
        Assert.Same(originalLocation, first.Location);
        Assert.Null(first.Location?.Page);
        Assert.Equal(captured.SnapshotHash, OfficeAiDocument.FromReadResult(new byte[] { 1 }, source).SnapshotHash);
    }

    [Fact]
    public async Task NamedTablesAndHeaderOnlyTablesKeepTheirIdentityAndSourceContext() {
        ReaderTable Table(string title, bool row) => new() { Title = title, Columns = new[] { "Item", "Count" },
            Rows = row ? new[] { new[] { "Pencil", "12" } } : Array.Empty<string[]>() };
        var source = new OfficeDocumentReadResult { Tables = new[] { Table("North", true), Table("South", true), Table("West", false) } };
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, source);
        Assert.Equal(5, document.Evidence.Count);
        Assert.Equal(new[] { "table", "table-row", "table", "table-row", "table" }, document.Evidence.Select(item => item.Kind));
        Assert.Equal("North\nItem | Count", document.Evidence[0].Text);
        Assert.Equal("North\nItem: Pencil | Count: 12", document.Evidence[1].Text);
        Assert.StartsWith("South\n", document.Evidence[3].Text);
        Assert.Equal("West\nItem | Count", document.Evidence[4].Text);
        Assert.Equal(new[] { "table-1", "table-1-row-1", "table-2", "table-2-row-1", "table-3" }, document.Evidence.Select(item => item.SourceBlockId));
        var executor = new Capture();
        await new OfficeAiEngine(executor).RunAsync(document, new() { Operation = OfficeAiOperation.Parse, Instruction = "Read all tables" });
        using var input = JsonDocument.Parse(Assert.Single(executor.Requests).InputJson);
        Assert.Equal(document.Evidence.Select(item => item.SourceBlockId), input.RootElement.GetProperty("evidence").EnumerateArray()
            .Select(item => item.GetProperty("sourceBlockId").GetString()));
        source.Tables[0].Title = "changed";
        Assert.StartsWith("North\n", document.Evidence[0].Text);
        Assert.NotEqual(document.SnapshotHash, OfficeAiDocument.FromReadResult(new byte[] { 1 }, source).SnapshotHash);
        Assert.Throws<InvalidDataException>(() => OfficeAiDocument.FromReadResult(new byte[] { 1 }, source, limits: new() { MaxDocumentBlocks = 4 }));
    }

    [Fact]
    public async Task EmptyPagesWithLocationOnlyRetainCoverageSelectionHashAndBounds() {
        var source = new OfficeDocumentReadResult {
            Pages = new[] {
                new OfficeDocumentPage { Number = 1, Blocks = new[] { new OfficeDocumentBlock { Text = "readable" } } },
                new OfficeDocumentPage { Location = new() { Page = 2 } }
            }
        };
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, source);
        Assert.Equal(new[] { 1, 2 }, document.Pages);
        var executor = new Capture();
        var engine = new OfficeAiEngine(executor);
        var result = await engine.RunAsync(document, new() { Instruction = "Read" });
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Equal(2, Assert.Single(result.EmptyPages));
        var selected = await engine.RunAsync(document, new() { Instruction = "Read", Pages = new[] { 2 } });
        Assert.Equal(2, Assert.Single(selected.EmptyPages));
        Assert.Single(executor.Requests);
        Assert.Throws<InvalidDataException>(() => OfficeAiDocument.FromReadResult(new byte[] { 1 }, source, limits: new() { MaxPages = 1 }));
        source.Pages = source.Pages.Take(1).ToArray();
        Assert.NotEqual(document.SnapshotHash, OfficeAiDocument.FromReadResult(new byte[] { 1 }, source).SnapshotHash);
    }

    [Fact]
    public async Task MixedReaderOwnersReachInferenceAndRetainPageProvenance() {
        var shared = new OfficeDocumentBlock { Id = "shared", Text = "aggregate" };
        var pageOnly = new OfficeDocumentBlock { Id = "page-only", Text = "page" };
        ReaderTable Table(string value) => new() { Columns = new[] { "Item" }, Rows = new[] { new[] { value } } };
        var aggregateTable = Table("aggregate table");
        var source = new OfficeDocumentReadResult {
            Blocks = new[] { shared }, Tables = new[] { aggregateTable },
            Pages = new[] { new OfficeDocumentPage { Number = 2, Blocks = new[] { shared, pageOnly }, Tables = new[] { aggregateTable, Table("page table") } } },
            Chunks = new[] { new ReaderChunk { Location = new() { Page = 3 }, Tables = new[] { Table("chunk table") } } }
        };
        Assert.Equal(2, source.EnumerateBlocks().Count());
        Assert.Equal(3, source.EnumerateTables().Count());
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, source);
        Assert.Equal(8, document.Evidence.Count);
        Assert.Equal(new int?[] { 2, 2, 2, 2, 2, 2, 3, 3 }, document.Evidence.Select(item => item.Page));
        Assert.Equal(new[] { 2, 3 }, document.Pages);
        var executor = new Capture();
        var result = await new OfficeAiEngine(executor).RunAsync(document, new() { Instruction = "Read all observations" });
        Assert.Equal(document.Evidence.Select(item => item.Id), result.ProcessedEvidenceIds);
        using var input = JsonDocument.Parse(Assert.Single(executor.Requests).InputJson);
        Assert.Equal(8, input.RootElement.GetProperty("evidence").GetArrayLength());
        pageOnly.Text = "changed";
        Assert.Contains(document.Evidence, item => item.Text == "page");
        Assert.NotEqual(document.SnapshotHash, OfficeAiDocument.FromReadResult(new byte[] { 1 }, source).SnapshotHash);
        Assert.Throws<InvalidDataException>(() => OfficeAiDocument.FromReadResult(new byte[] { 1 }, source,
            limits: new() { MaxDocumentBlocks = 4 }));
    }

    [Fact]
    public async Task MultipleImagesOnOnePageUseTheirOwnCountBudget() {
        var images = new[] {
            new OfficeAiImage("first", 1, "image/png", new byte[] { 1 }, 1, 1),
            new OfficeAiImage("second", 1, "image/png", new byte[] { 2 }, 1, 1)
        };
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, new(), images,
            new() { MaxPages = 1, MaxDocumentImages = 2 });
        Assert.Equal(2, document.Images.Count);
        Assert.Equal(1, Assert.Single(document.Pages));
        var executor = new Capture();
        await Assert.ThrowsAsync<ArgumentException>(() => new OfficeAiEngine(executor).RunAsync(document,
            new() { Instruction = "Read", Limits = new() { MaxDocumentImages = 1 } }));
        Assert.Empty(executor.Requests);
        Assert.Throws<ArgumentException>(() => OfficeAiDocument.FromReadResult(new byte[] { 1 }, new(), images,
            new() { MaxPages = 1, MaxDocumentImages = 1 }));
    }

    private sealed class Capture : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile { get; } = new() { Id = "fixture", Provider = "fixture", Model = "fixture", IsLocal = true };
        public List<OfficeAiExecutionRequest> Requests { get; } = new();
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            Requests.Add(request);
            return Task.FromResult(new OfficeAiExecutionResponse("{\"status\":\"insufficient\",\"claims\":[],\"fields\":[],\"blocks\":[],\"tables\":[]}"));
        }
    }
}
