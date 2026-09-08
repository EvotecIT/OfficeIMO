using System.Text.Json;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class SnapshotCoverageTests {
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
        Assert.Equal(5, document.Evidence.Count);
        Assert.Equal(new int?[] { 2, 2, 2, 2, 3 }, document.Evidence.Select(item => item.Page));
        Assert.Equal(new[] { 2, 3 }, document.Pages);
        var executor = new Capture();
        var result = await new OfficeAiEngine(executor).RunAsync(document, new() { Instruction = "Read all observations" });
        Assert.Equal(document.Evidence.Select(item => item.Id), result.ProcessedEvidenceIds);
        using var input = JsonDocument.Parse(Assert.Single(executor.Requests).InputJson);
        Assert.Equal(5, input.RootElement.GetProperty("evidence").GetArrayLength());
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
