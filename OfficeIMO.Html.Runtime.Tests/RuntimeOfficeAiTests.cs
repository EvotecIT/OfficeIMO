using System.Text.Json;
using global::OfficeIMO.AI;
using global::OfficeIMO.AI.Html;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Conformance;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeOfficeAiTests {
    [Fact]
    public async Task OfficeAiExecutorDrivesTheBoundedPlannerWithoutEnteringRuntimeCore() {
        var executor = new ScriptedExecutor();
        var adapter = new HtmlAutomationAiPlanner(executor, turn => $"Complete the approved workflow at step {turn.Step}.");
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
        Assert.Equal(4, executor.Requests.Count);
        Assert.All(executor.Requests, request => {
            Assert.Contains(HtmlAutomationToolNames.Act, request.OutputSchema, StringComparison.Ordinal);
            Assert.Contains(HtmlAutomationToolNames.Capture, request.OutputSchema, StringComparison.Ordinal);
            Assert.Contains("\"observation\"", request.InputJson, StringComparison.Ordinal);
            Assert.Empty(request.Images);
        });
        Assert.Equal(HtmlAutomationToolCatalog.GetDefinitions().Select(item => item.InputSchema.GetRawText()),
            adapter.Tools.Select(item => item.InputSchema.GetRawText()));
        HtmlRuntimeQualificationManifest manifest = HtmlRuntimeQualificationCatalog.Get("programmatic-automation-v1");
        HtmlRuntimeConsumerExpectation evidence = manifest.Consumers.Single();
        Assert.Equal("officeimo-ai-html", evidence.Id);
        Assert.Equal("OfficeIMO.AI.IOfficeAiExecutor", evidence.Contract);
        HtmlRuntimeConsumerQualificationResult consumer = HtmlRuntimeQualificationCatalog.EvaluateConsumer(manifest, evidence.Id,
            passedWorkflows: 1, failedWorkflows: 0);
        HtmlRuntimeConformanceReport providerReport = await HtmlRuntimeConformanceSuite.RunAsync(runtime, manifest);
        HtmlRuntimeProfileQualificationResult qualification = HtmlRuntimeQualificationCatalog.Evaluate(manifest, providerReport, consumer);
        Assert.True(consumer.Passed);
        Assert.True(qualification.Passed);
    }

    [Fact]
    public async Task OfficeAiPlannerRejectsOversizedToolArguments() {
        var executor = new SingleResponseExecutor(Decision(false,
            Call("oversized", HtmlAutomationToolNames.Act,
                "{\"action\":\"Fill\",\"css\":\"#name\",\"value\":\"" + new string('x', 256) + "\"}")));
        var adapter = new HtmlAutomationAiPlanner(executor, _ => "bounded",
            new HtmlAutomationAiPlannerOptions { MaxArgumentBytes = 64 });

        InvalidOperationException failure = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            adapter.CreatePlanner()(Turn(), CancellationToken.None));

        Assert.Contains("byte limit", failure.Message, StringComparison.Ordinal);
        Assert.Single(executor.Requests);
    }

    [Fact]
    public async Task OfficeAiPlannerRejectsUndeclaredToolsAndDuplicateProperties() {
        var unknown = new SingleResponseExecutor(Decision(false, Call("bad", "unknown", "{}")));
        var unknownAdapter = new HtmlAutomationAiPlanner(unknown, _ => "bounded");
        InvalidOperationException unknownFailure = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            unknownAdapter.CreatePlanner()(Turn(), CancellationToken.None));
        Assert.Contains("undeclared tool", unknownFailure.Message, StringComparison.Ordinal);

        var duplicate = new SingleResponseExecutor("{\"isComplete\":false,\"isComplete\":true,\"calls\":[]}");
        var duplicateAdapter = new HtmlAutomationAiPlanner(duplicate, _ => "bounded");
        InvalidOperationException duplicateFailure = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            duplicateAdapter.CreatePlanner()(Turn(), CancellationToken.None));
        Assert.Contains("duplicate", duplicateFailure.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static HtmlAutomationTurn Turn() => new() {
        Observation = new HtmlPageObservation { ProviderId = "test", ContextId = "test", PageId = "test" }
    };

    private static string Decision(bool complete, params string[] calls) =>
        $"{{\"isComplete\":{complete.ToString().ToLowerInvariant()},\"message\":{(complete ? "\"accepted\"" : "null")},\"calls\":[{string.Join(',', calls)}]}}";

    private static string Call(string id, string name, string arguments) =>
        $"{{\"id\":{JsonSerializer.Serialize(id)},\"name\":{JsonSerializer.Serialize(name)},\"arguments\":{arguments}}}";

    private sealed class ScriptedExecutor : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile { get; } = ProfileFor("scripted");
        public List<OfficeAiExecutionRequest> Requests { get; } = new();

        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            Requests.Add(request);
            string response = Requests.Count switch {
                1 => Decision(false, Call("fill", HtmlAutomationToolNames.Act, "{\"css\":\"#name\",\"action\":\"Fill\",\"value\":\"Ada\"}")),
                2 => Decision(false, Call("submit", HtmlAutomationToolNames.Act, "{\"css\":\"#submit\",\"action\":\"Click\",\"waitForReady\":false}")),
                3 => Decision(false,
                    Call("wait", HtmlAutomationToolNames.Act, "{\"css\":\"output\",\"action\":\"Wait\",\"waitState\":\"Text\",\"value\":\"Accepted Ada\"}"),
                    Call("capture", HtmlAutomationToolNames.Capture, "{}")),
                _ => Decision(true)
            };
            return Task.FromResult(new OfficeAiExecutionResponse(response));
        }
    }

    private sealed class SingleResponseExecutor(string response) : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile { get; } = ProfileFor("single");
        public List<OfficeAiExecutionRequest> Requests { get; } = new();
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            Requests.Add(request);
            return Task.FromResult(new OfficeAiExecutionResponse(response));
        }
    }

    private static OfficeAiExecutionProfile ProfileFor(string id) => new() {
        Id = id,
        Provider = "test",
        Model = "test",
        IsLocal = true,
        EnforcesJsonSchema = true,
        MaxRequestCharacters = 2_000_000
    };
}
