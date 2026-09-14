using System.Text.Json;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class ToolPlanningTests {
    [Fact]
    public async Task PlannerBuildsOwnedSchemaAndReturnsDetachedBoundedCalls() {
        using JsonDocument schema = JsonDocument.Parse("""{"type":"object","additionalProperties":false,"required":["value"],"properties":{"value":{"type":"string"}}}""");
        var executor = new RecordingExecutor("""{"isComplete":false,"message":null,"calls":[{"id":"set-1","name":"set_value","arguments":{"value":"ready"}}]}""");
        var planner = new OfficeAiToolPlanner(executor);

        OfficeAiToolPlanningDecision decision = await planner.PlanAsync(new OfficeAiToolPlanningRequest {
            RequestId = "tool-test",
            Instructions = "Choose the next declared operation.",
            InputJson = "{\"state\":\"pending\"}",
            Tools = new[] { new OfficeAiToolDefinition("set_value", "Set a value.", schema.RootElement) },
            MaxToolCalls = 1
        });

        Assert.False(decision.IsComplete);
        OfficeAiToolCall call = Assert.Single(decision.Calls);
        Assert.Equal("set-1", call.Id);
        Assert.Equal("set_value", call.Name);
        Assert.Equal("ready", call.Arguments.GetProperty("value").GetString());
        Assert.Contains("\"const\":\"set_value\"", executor.Request!.OutputSchema, StringComparison.Ordinal);
        Assert.DoesNotContain("\"oneOf\"", executor.Request.OutputSchema, StringComparison.Ordinal);
        using JsonDocument outputSchema = JsonDocument.Parse(executor.Request.OutputSchema);
        Assert.Equal(new[] { "isComplete", "message", "calls" }, outputSchema.RootElement.GetProperty("required")
            .EnumerateArray().Select(item => item.GetString()));
        Assert.Empty(executor.Request.Images);
    }

    [Fact]
    public async Task PlannerRejectsTruncatedAndStructurallyInvalidDecisions() {
        using JsonDocument schema = JsonDocument.Parse("{\"type\":\"object\",\"additionalProperties\":false,\"properties\":{}}");
        OfficeAiToolPlanningRequest request = new() {
            RequestId = "tool-test",
            Instructions = "Choose a declared operation.",
            InputJson = "{}",
            Tools = new[] { new OfficeAiToolDefinition("run", "Run.", schema.RootElement) }
        };
        var truncated = new OfficeAiToolPlanner(new RecordingExecutor("{}", isComplete: false));
        await Assert.ThrowsAsync<InvalidOperationException>(() => truncated.PlanAsync(request));

        var malformed = new OfficeAiToolPlanner(new RecordingExecutor(
            "{\"isComplete\":false,\"calls\":[{\"id\":\"one\",\"name\":\"missing\",\"arguments\":{}}]}"));
        InvalidOperationException failure = await Assert.ThrowsAsync<InvalidOperationException>(() => malformed.PlanAsync(request));
        Assert.Contains("undeclared tool", failure.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("{\"action\":\"Inspect\",\"action\":\"Click\"}", "duplicate")]
    [InlineData("{\"action\":\"Inspect\",\"unexpected\":true}", "unknown")]
    [InlineData("{}", "missing required")]
    public async Task PlannerRejectsArgumentsThatDoNotMatchTheDeclaredSchema(string arguments, string message) {
        using JsonDocument schema = JsonDocument.Parse("""{"type":"object","additionalProperties":false,"required":["action"],"properties":{"action":{"type":"string","enum":["Inspect","Click"]}}}""");
        var planner = new OfficeAiToolPlanner(new RecordingExecutor(
            $"{{\"isComplete\":false,\"message\":null,\"calls\":[{{\"id\":\"one\",\"name\":\"act\",\"arguments\":{arguments}}}]}}"));
        var request = new OfficeAiToolPlanningRequest {
            RequestId = "tool-validation", Instructions = "Choose an action.", InputJson = "{}",
            Tools = new[] { new OfficeAiToolDefinition("act", "Act.", schema.RootElement) }
        };

        InvalidOperationException failure = await Assert.ThrowsAsync<InvalidOperationException>(() => planner.PlanAsync(request));

        Assert.Contains(message, failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task PlannerRemovesStrictProviderNullsForOptionalArguments() {
        using JsonDocument schema = JsonDocument.Parse("""{"type":"object","additionalProperties":false,"required":["action"],"properties":{"action":{"type":"string"},"value":{"type":"string"}}}""");
        var executor = new RecordingExecutor("""{"isComplete":false,"message":null,"calls":[{"id":"one","name":"act","arguments":{"action":"Inspect","value":null}}]}""");

        OfficeAiToolPlanningDecision decision = await new OfficeAiToolPlanner(executor).PlanAsync(new OfficeAiToolPlanningRequest {
            RequestId = "strict-null", Instructions = "Choose an action.", InputJson = "{}",
            Tools = new[] { new OfficeAiToolDefinition("act", "Act.", schema.RootElement) }
        });

        JsonElement arguments = Assert.Single(decision.Calls).Arguments;
        Assert.Equal("Inspect", arguments.GetProperty("action").GetString());
        Assert.False(arguments.TryGetProperty("value", out _));
        Assert.Contains("\"type\":\"null\"", executor.Request!.OutputSchema, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("{\"type\":\"object\",\"properties\":{}}")]
    [InlineData("{\"type\":\"object\",\"additionalProperties\":false,\"properties\":{},\"patternProperties\":{}}")]
    public void ToolDefinitionRejectsOpenOrUnsupportedSchemas(string schemaJson) {
        using JsonDocument schema = JsonDocument.Parse(schemaJson);

        Assert.Throws<ArgumentException>(() => new OfficeAiToolDefinition("unsafe", "Unsafe.", schema.RootElement));
    }

    private sealed class RecordingExecutor(string response, bool isComplete = true) : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile { get; } = new() {
            Id = "tool-tests",
            Provider = "test",
            Model = "test",
            IsLocal = true,
            EnforcesJsonSchema = true,
            MaxRequestCharacters = 2_000_000
        };

        public OfficeAiExecutionRequest? Request { get; private set; }

        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request,
            CancellationToken cancellationToken = default) {
            Request = request;
            return Task.FromResult(new OfficeAiExecutionResponse(response, IsComplete: isComplete));
        }
    }
}
