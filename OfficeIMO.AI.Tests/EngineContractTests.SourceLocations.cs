using System.Text.Json;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed partial class EngineContractTests {
    [Fact]
    public async Task CitationsRetainImmutableSheetAndTableRowCoordinates() {
        var location = new ReaderLocation { Path = "book.xlsx", Sheet = "Invoices", A1Range = "A2:B3", TableIndex = 0, SourceBlockIndex = 2 };
        var result = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Excel, Source = new() { Path = "book.xlsx" },
            Tables = [new() { Columns = ["Name", "Amount"], Rows = [["Alpha", "42"], ["Beta", "7"]], Location = location }]
        };
        var document = OfficeAiDocument.FromReadResult([1], result);
        location.Path = "mutated.xlsx"; location.Sheet = "Mutated"; location.A1Range = "Z99";
        var executor = new Executor(Field("42", "e2"));
        var answer = await new OfficeAiEngine(executor).RunAsync(document, Request() with {
            Operation = OfficeAiOperation.ExtractFields, Fields = [new("amount", OfficeAiFieldType.Integer)]
        });
        var citation = Assert.Single(Assert.Single(answer.Fields).Citations);
        Assert.Equal("book.xlsx", citation.SourceLocation!.Path);
        Assert.Equal("Invoices", citation.SourceLocation.Sheet);
        Assert.Equal("A2:B3", citation.SourceLocation.A1Range);
        Assert.Equal(0, citation.SourceLocation.TableIndex);
        Assert.Equal(1, citation.SourceLocation.TableRowNumber);
        using var request = JsonDocument.Parse(Assert.Single(executor.Requests).InputJson);
        var source = request.RootElement.GetProperty("evidence")[1].GetProperty("sourceLocation");
        Assert.Equal("Invoices", source.GetProperty("Sheet").GetString());
        Assert.Equal(1, source.GetProperty("TableRowNumber").GetInt32());
        var restored = JsonSerializer.Deserialize<OfficeAiResult>(JsonSerializer.Serialize(answer))!;
        Assert.Equal(citation.SourceLocation, Assert.Single(Assert.Single(restored.Fields).Citations).SourceLocation);
    }

    [Fact]
    public async Task SamePageNumbersRemainDistinguishableAcrossContainerSources() {
        var source = new OfficeDocumentReadResult { Source = new() { Path = "mail.eml" }, Blocks = [
            new() { Id = "a", Text = "Amount 42", Location = new() { Path = "mail.eml::a.pdf", Page = 1 } },
            new() { Id = "b", Text = "Amount 7", Location = new() { Path = "mail.eml::b.pdf", Page = 1 } }
        ] };
        var document = OfficeAiDocument.FromReadResult([1], source);
        var executor = new Executor("{\"claims\":[{\"text\":\"The amounts differ.\",\"evidence\":[{\"id\":\"e1\",\"quote\":\"42\"},{\"id\":\"e2\",\"quote\":\"7\"}]}],\"fields\":[],\"blocks\":[],\"tables\":[]}");
        var result = await new OfficeAiEngine(executor).RunAsync(document, Request());
        var citations = Assert.Single(result.Claims).Citations;
        Assert.Equal(new[] { "mail.eml::a.pdf", "mail.eml::b.pdf" }, citations.Select(item => item.SourceLocation!.Path));
        Assert.All(citations, item => Assert.Equal(1, item.Page));
        source.Blocks[0].Location.Path = "mail.eml::other.pdf";
        Assert.NotEqual(document.SnapshotHash, OfficeAiDocument.FromReadResult([1], source).SnapshotHash);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ParseCoordinatesRequireACommonSourceContainer(bool sameSource) {
        const string firstPath = "mail.eml::a.pdf";
        var source = new OfficeDocumentReadResult { Blocks = [
            new() { Id = "a", Text = "Amount 42", Location = new() { Path = firstPath, Page = 1 } },
            new() { Id = "b", Text = "Amount 7", Location = new() { Path = sameSource ? firstPath : "mail.eml::b.pdf", Page = 1 } }
        ] };
        var evidence = new[] { new { id = "e1", quote = "42" }, new { id = "e2", quote = "7" } };
        string response = JsonSerializer.Serialize(new {
            claims = Array.Empty<object>(), fields = Array.Empty<object>(),
            blocks = new[] { new { kind = "paragraph", text = "Amounts 42 and 7", evidence } },
            tables = new[] { new { title = "Amounts", columns = new[] { "Amount" }, rows = new[] { new[] { "42" }, new[] { "7" } }, evidence } }
        });
        var document = OfficeAiDocument.FromReadResult([1], source);
        var result = await new OfficeAiEngine(new Executor(response)).RunAsync(document, Request() with { Operation = OfficeAiOperation.Parse });
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        ReaderLocation[] locations = [Assert.Single(result.Blocks).Block.Location,
            Assert.IsType<ReaderLocation>(Assert.Single(result.Tables).Table.Location)];
        Assert.All(locations, location => {
            Assert.Equal(sameSource ? firstPath : null, location.Path);
            Assert.Equal(sameSource ? (int?)1 : null, location.Page);
        });
    }
}
