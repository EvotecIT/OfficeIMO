using System.Runtime.CompilerServices;
using System.Text.Json;
using Microsoft.Extensions.AI;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.MicrosoftExtensionsAI;
using OfficeIMO.Html.Runtime.Conformance;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeMicrosoftExtensionsAiTests {
    [Fact]
    public async Task ChatClientDrivesTheBoundedPlannerWithoutEnteringRuntimeCore() {
        using var client = new ScriptedChatClient();
        var adapter = new HtmlAutomationChatPlanner(client, turn => new HtmlAutomationChatRequest {
            Messages = new[] { new ChatMessage(ChatRole.User, $"step={turn.Step}\n{HtmlRuntimeJson.Serialize(turn.Observation)}") },
            Options = new ChatOptions { ModelId = "caller-selected-model" }
        });
        var runtime = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
        await using IHtmlRuntimeContext context = await runtime.CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = "<label>Name <input id='name'></label><button id='submit'>Submit</button><output>Pending</output>",
            Scripts = new[] { "document.querySelector('#submit').onclick=()=>setTimeout(()=>document.querySelector('output').textContent='Accepted '+document.querySelector('#name').value,25)" }
        });

        HtmlAutomationRunResult result = await new HtmlAutomationRunner().RunAsync(page, adapter.CreatePlanner(),
            new HtmlAutomationRunOptions { MaxSteps = 5, MaxCallsPerStep = 2 });

        Assert.True(result.IsComplete);
        Assert.Equal("accepted", result.Message);
        Assert.Equal(4, result.ToolResults.Count);
        Assert.Equal("Accepted Ada", result.ToolResults.Single(item => item.CallId == "capture").Capture!.Document.QuerySelector("output")!.TextContent);
        Assert.Equal(4, client.Requests);
        Assert.All(client.Options, options => {
            Assert.Equal("caller-selected-model", options.ModelId);
            Assert.Equal(HtmlAutomationToolCatalog.GetDefinitions().Select(item => item.Name),
                options.Tools!.OfType<AIFunctionDeclaration>().Select(item => item.Name));
        });
        Assert.Equal(HtmlAutomationToolCatalog.GetDefinitions().Select(item => item.InputSchema.GetRawText()),
            adapter.Tools.Cast<AIFunctionDeclaration>().Select(item => item.JsonSchema.GetRawText()));
        HtmlRuntimeQualificationManifest manifest = HtmlRuntimeQualificationCatalog.Get("programmatic-automation-v1");
        HtmlRuntimeConsumerExpectation evidence = manifest.Consumers.Single();
        Assert.Equal("microsoft-extensions-ai", evidence.Id);
        Assert.Equal("10.10.0", evidence.PackageVersion);
        HtmlRuntimeConsumerQualificationResult consumer = HtmlRuntimeQualificationCatalog.EvaluateConsumer(manifest, evidence.Id,
            passedWorkflows: 1, failedWorkflows: 0);
        HtmlRuntimeConformanceReport providerReport = await HtmlRuntimeConformanceSuite.RunAsync(runtime, manifest);
        HtmlRuntimeProfileQualificationResult qualification = HtmlRuntimeQualificationCatalog.Evaluate(manifest, providerReport, consumer);
        Assert.True(consumer.Passed);
        Assert.True(qualification.Passed);
    }

    [Fact]
    public async Task ChatPlannerRejectsReservedToolConflictsBeforeCallingTheClient() {
        using var client = new SingleResponseChatClient(new ChatResponse(new ChatMessage(ChatRole.Assistant, "unused")));
        using JsonDocument schema = JsonDocument.Parse("{}");
        var adapter = new HtmlAutomationChatPlanner(client, _ => new HtmlAutomationChatRequest {
            Messages = new[] { new ChatMessage(ChatRole.User, "conflict") },
            Options = new ChatOptions { Tools = new[] {
                (AITool)AIFunctionFactory.CreateDeclaration(HtmlAutomationToolNames.Act, "conflict", schema.RootElement)
            } }
        });

        InvalidOperationException failure = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            adapter.CreatePlanner()(Turn(), CancellationToken.None));

        Assert.Contains("conflicting reserved OfficeIMO tool", failure.Message, StringComparison.Ordinal);
        Assert.Equal(0, client.Requests);
    }

    [Fact]
    public async Task ChatPlannerRejectsOversizedToolArgumentsBeforeReturningADecision() {
        ChatMessage response = Calls(Call("oversized", HtmlAutomationToolNames.Act,
            ("action", "Fill"), ("css", "#name"), ("value", new string('x', 256))));
        using var client = new SingleResponseChatClient(new ChatResponse(response));
        var adapter = new HtmlAutomationChatPlanner(client, _ => new HtmlAutomationChatRequest {
            Messages = new[] { new ChatMessage(ChatRole.User, "bounded") }
        }, new HtmlAutomationChatPlannerOptions { MaxArgumentBytes = 64 });

        InvalidOperationException failure = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            adapter.CreatePlanner()(Turn(), CancellationToken.None));

        Assert.Contains("byte limit", failure.Message, StringComparison.Ordinal);
        Assert.Equal(1, client.Requests);
    }

    [Fact]
    public async Task ChatPlannerRejectsFunctionCallsThatFailedDuringProviderMapping() {
        FunctionCallContent malformed = Call("malformed", HtmlAutomationToolNames.Capture);
        malformed.Exception = new JsonException("provider could not map arguments");
        using var client = new SingleResponseChatClient(new ChatResponse(Calls(malformed)));
        var adapter = new HtmlAutomationChatPlanner(client, _ => new HtmlAutomationChatRequest {
            Messages = new[] { new ChatMessage(ChatRole.User, "malformed") }
        });

        InvalidOperationException failure = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            adapter.CreatePlanner()(Turn(), CancellationToken.None));

        Assert.Contains("invalid OfficeIMO tool call", failure.Message, StringComparison.Ordinal);
        Assert.IsType<JsonException>(failure.InnerException);
        Assert.Equal(1, client.Requests);
    }

    private static HtmlAutomationTurn Turn() => new() {
        Observation = new HtmlPageObservation { ProviderId = "test", ContextId = "test", PageId = "test" }
    };

    private static ChatMessage Calls(params FunctionCallContent[] calls) =>
        new(ChatRole.Assistant, calls.Cast<AIContent>().ToList());

    private static FunctionCallContent Call(string id, string name, params (string Name, object? Value)[] arguments) =>
        new(id, name, arguments.ToDictionary(item => item.Name, item => item.Value));

    private sealed class ScriptedChatClient : IChatClient {
        internal int Requests { get; private set; }
        internal List<ChatOptions> Options { get; } = new();

        public Task<ChatResponse> GetResponseAsync(IEnumerable<ChatMessage> messages, ChatOptions? options = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            Requests++;
            Options.Add(options?.Clone() ?? new ChatOptions());
            ChatMessage response = Requests switch {
                1 => Calls(Call("fill", HtmlAutomationToolNames.Act, ("css", "#name"), ("action", "Fill"), ("value", "Ada"))),
                2 => Calls(Call("submit", HtmlAutomationToolNames.Act, ("css", "#submit"), ("action", "Click"), ("waitForReady", false))),
                3 => Calls(
                    Call("wait", HtmlAutomationToolNames.Act, ("css", "output"), ("action", "Wait"), ("waitState", "Text"), ("value", "Accepted Ada")),
                    Call("capture", HtmlAutomationToolNames.Capture)),
                _ => new ChatMessage(ChatRole.Assistant, "accepted")
            };
            return Task.FromResult(new ChatResponse(response));
        }

        public async IAsyncEnumerable<ChatResponseUpdate> GetStreamingResponseAsync(IEnumerable<ChatMessage> messages,
            ChatOptions? options = null, [EnumeratorCancellation] CancellationToken cancellationToken = default) {
            await Task.CompletedTask;
            yield break;
        }

        public object? GetService(Type serviceType, object? serviceKey = null) => null;
        public void Dispose() { }

    }

    private sealed class SingleResponseChatClient(ChatResponse response) : IChatClient {
        internal int Requests { get; private set; }
        public Task<ChatResponse> GetResponseAsync(IEnumerable<ChatMessage> messages, ChatOptions? options = null,
            CancellationToken cancellationToken = default) {
            Requests++;
            return Task.FromResult(response);
        }
        public async IAsyncEnumerable<ChatResponseUpdate> GetStreamingResponseAsync(IEnumerable<ChatMessage> messages,
            ChatOptions? options = null, [EnumeratorCancellation] CancellationToken cancellationToken = default) {
            await Task.CompletedTask;
            yield break;
        }
        public object? GetService(Type serviceType, object? serviceKey = null) => null;
        public void Dispose() { }
    }
}
