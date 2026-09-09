using System.Text;
using System.Text.Json;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class ExtractionSchemaTests {
    [Theory]
    [InlineData(OfficeAiOperation.Ask)]
    [InlineData(OfficeAiOperation.Explain)]
    [InlineData(OfficeAiOperation.Summarize)]
    [InlineData(OfficeAiOperation.Parse)]
    [InlineData(OfficeAiOperation.ExtractFields)]
    public async Task EmptyValidatedOutputDerivesInsufficientEvidence(OfficeAiOperation operation) {
        string fields = operation == OfficeAiOperation.ExtractFields
            ? "{\"field1\":{\"status\":\"missing\",\"rawValue\":null,\"evidence\":[]}}" : "[]";
        var executor = new Executor("{\"claims\":[],\"fields\":" + fields + ",\"blocks\":[],\"tables\":[]}");
        var document = OfficeAiDocument.FromReadResult(Encoding.UTF8.GetBytes("42"), new OfficeDocumentReadResult {
            Blocks = new[] { new OfficeDocumentBlock { Text = "42" } }
        });
        var result = await new OfficeAiEngine(executor).RunAsync(document, new() { Operation = operation,
            Instruction = "Find the requested information.",
            Fields = operation == OfficeAiOperation.ExtractFields ? new[] { new OfficeAiFieldDefinition("amount") } : Array.Empty<OfficeAiFieldDefinition>() });
        Assert.Equal(OfficeAiResultStatus.InsufficientEvidence, result.Status);
        Assert.Empty(result.Claims);
        Assert.Empty(result.Tables);
        if (operation == OfficeAiOperation.ExtractFields) Assert.Equal(OfficeAiFieldStatus.Missing, Assert.Single(result.Fields).Status);
    }

    [Fact]
    public async Task RequestedKeysPreserveNamesAndResultOrder() {
        string[] names = { "amount/a~b", "reference\"id", "总额" };
        var fields = names.Select((name,index) => "field" + (index + 1)).Reverse().ToDictionary(key => key,
            _ => (object)new { status = "present", rawValue = "42", evidence = new[] { new { id = "e1", quote = "42" } } });
        var executor = new Executor(JsonSerializer.Serialize(new {
            claims = Array.Empty<object>(), fields, blocks = Array.Empty<object>(), tables = Array.Empty<object>()
        }));
        var result = await Run(executor, names);
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        Assert.Equal(names, result.Fields.Select(field => field.Name));
        using var schema = JsonDocument.Parse(executor.Schema!);
        var shape = schema.RootElement.GetProperty("properties").GetProperty("fields");
        Assert.Equal(new[] { "field1", "field2", "field3" }, shape.GetProperty("required").EnumerateArray().Select(item => item.GetString()));
        Assert.Equal(new[] { "field1", "field2", "field3" }, shape.GetProperty("properties").EnumerateObject().Select(item => item.Name));
        Assert.False(shape.GetProperty("additionalProperties").GetBoolean());
        using var input = JsonDocument.Parse(executor.Input!);
        var requested = input.RootElement.GetProperty("fields").EnumerateArray().ToArray();
        Assert.Equal(names, requested.Select(item => item.GetProperty("name").GetString()));
        Assert.Equal(new[] { "field1", "field2", "field3" }, requested.Select(item => item.GetProperty("key").GetString()));
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("{\"unrequested\":{\"status\":\"missing\",\"rawValue\":null,\"evidence\":[]}}")]
    [InlineData("{\"field1\":{\"status\":\"missing\",\"rawValue\":null,\"evidence\":[]},\"field1\":{\"status\":\"missing\",\"rawValue\":null,\"evidence\":[]}}")]
    public async Task MissingUnknownAndDuplicateKeysAreRejectedByLocalValidation(string fields) {
        var executor = new Executor("{\"claims\":[],\"fields\":" + fields + ",\"blocks\":[],\"tables\":[]}");
        Assert.Equal(OfficeAiResultStatus.InvalidResponse, (await Run(executor, new[] { "amount", "reference" })).Status);
    }

    [Theory]
    [InlineData("present", "42", "[{\"id\":\"e1\",\"quote\":\"42\"}]", true)]
    [InlineData("present", null, "[{\"id\":\"e1\",\"quote\":\"42\"}]", false)]
    [InlineData("present", "42", "[]", false)]
    [InlineData("missing", null, "[]", true)]
    [InlineData("missing", "42", "[]", false)]
    [InlineData("missing", null, "[{\"id\":\"e1\",\"quote\":\"42\"}]", false)]
    [InlineData("ambiguous", null, "[{\"id\":\"e1\",\"quote\":\"42\"}]", true)]
    [InlineData("conflicting", null, "[{\"id\":\"e1\",\"quote\":\"42\"}]", true)]
    [InlineData("ambiguous", null, "[]", false)]
    public async Task FieldStatesRetainTheirValueAndEvidenceContracts(string status, string? raw, string evidence, bool valid) {
        string value = JsonSerializer.Serialize(raw);
        var executor = new Executor("{\"claims\":[],\"fields\":{\"field1\":{\"status\":\"" + status
            + "\",\"rawValue\":" + value + ",\"evidence\":" + evidence + "}},\"blocks\":[],\"tables\":[]}");
        var result = await Run(executor, new[] { "amount" });
        Assert.Equal(valid, result.Status != OfficeAiResultStatus.InvalidResponse);
        if (valid) Assert.Single(result.Fields);
    }

    private static Task<OfficeAiResult> Run(Executor executor, string[] names) {
        var document = OfficeAiDocument.FromReadResult(Encoding.UTF8.GetBytes("42"), new OfficeDocumentReadResult {
            Blocks = new[] { new OfficeDocumentBlock { Text = "42" } }
        });
        return new OfficeAiEngine(executor).RunAsync(document, new() { Operation = OfficeAiOperation.ExtractFields,
            Instruction = "Extract the requested fields.", Fields = names.Select(name => new OfficeAiFieldDefinition(name)).ToArray() });
    }

    private sealed class Executor(string response) : IOfficeAiExecutor {
        public string? Schema;
        public string? Input;
        public OfficeAiExecutionProfile Profile { get; } = new() { Id = "fields", Provider = "fixture", Model = "fixture", IsLocal = true };
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            Schema = request.OutputSchema;
            Input = request.InputJson;
            return Task.FromResult(new OfficeAiExecutionResponse(response));
        }
    }
}
