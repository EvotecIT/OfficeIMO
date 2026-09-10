using System.Text.Json;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed partial class EngineContractTests {
    [Fact]
    public async Task PriorDiscussionCannotBecomeCitableSourceEvidence() {
        var executor = new Executor("""{"claims":[{"text":"The prior answer says invented","evidence":[{"id":"e1","quote":"invented"}]}],"fields":[],"blocks":[],"tables":[]}""");
        var result = await new OfficeAiEngine(executor).RunAsync(Document("actual source"), Request() with {
            ConversationContext = "Previous answer: invented"
        });
        Assert.Empty(result.Claims);
        Assert.Equal(OfficeAiResultStatus.InvalidResponse, result.Status);
        using var input = JsonDocument.Parse(Assert.Single(executor.Requests).InputJson);
        Assert.Equal("Previous answer: invented", input.RootElement.GetProperty("conversationContext").GetString());
        Assert.Equal("actual source", input.RootElement.GetProperty("evidence")[0].GetProperty("text").GetString());
    }

    [Fact]
    public async Task OversizedPriorDiscussionIsRejectedBeforeExecution() {
        var executor = new Executor(Empty);
        await Assert.ThrowsAsync<ArgumentException>(() => new OfficeAiEngine(executor).RunAsync(Document("source"), Request() with {
            ConversationContext = new string('x', 8001)
        }));
        Assert.Empty(executor.Requests);
    }
}
