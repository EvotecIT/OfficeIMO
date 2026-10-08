using System.Text.Json;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeAutomationToolTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TypedObserveArgumentsPreserveScreenshotRequestDuringNormalization(bool screenshot) {
        HtmlAutomationToolCall call = HtmlAutomationToolCall.Observe("observe", new HtmlPageObservationRequest {
            IncludeScreenshotReference = screenshot
        });

        JsonElement arguments = HtmlAutomationToolCatalog.NormalizeArguments(call.Name, call.Arguments);

        Assert.Equal(screenshot, arguments.GetProperty("includeScreenshotReference").GetBoolean());
    }

    [Fact]
    public async Task TypedObserveDispatchesAndReportsUnsupportedScreenshotCapability() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() { Html = "<p>Observed</p>" });
        var dispatcher = new HtmlAutomationToolDispatcher();

        HtmlAutomationToolResult observed = await dispatcher.ExecuteAsync(page, HtmlAutomationToolCall.Observe("observe"));
        HtmlAutomationToolResult screenshot = await dispatcher.ExecuteAsync(page, HtmlAutomationToolCall.Observe("screenshot", new() {
            IncludeScreenshotReference = true
        }));

        Assert.True(observed.IsSuccess);
        Assert.NotNull(observed.Observation);
        Assert.False(screenshot.IsSuccess);
        Assert.Contains("screenshot", screenshot.Error!, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("unknown", screenshot.Error!, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ToolCatalogContainsDetachedJsonSchemasWithoutAnAgentSdkDependency() {
        IReadOnlyList<HtmlAutomationToolDefinition> tools = HtmlAutomationToolCatalog.GetDefinitions();

        Assert.Equal(4, tools.Count);
        Assert.Equal(tools.Count, tools.Select(tool => tool.Name).Distinct(StringComparer.Ordinal).Count());
        Assert.All(tools, tool => {
            Assert.Equal(JsonValueKind.Object, tool.InputSchema.ValueKind);
            Assert.Equal("object", tool.InputSchema.GetProperty("type").GetString());
        });
    }

    [Fact]
    public async Task DispatcherUsesTheSameRevisionBoundActionContract() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = "<button>Increment</button><output>0</output>",
            Scripts = new[] { "document.querySelector('button').onclick=()=>document.querySelector('output').textContent='1'" }
        });
        var dispatcher = new HtmlAutomationToolDispatcher();
        HtmlObservedElementReference reference = (await page.ObserveAsync(new() { ActionableOnly = true })).Elements.Single().Reference;

        HtmlAutomationToolResult first = await dispatcher.ExecuteAsync(page, HtmlAutomationToolCall.Act("first", new() {
            Reference = reference, Action = HtmlAutomationAction.Click, WaitForReady = false
        }));
        HtmlAutomationToolResult second = await dispatcher.ExecuteAsync(page, HtmlAutomationToolCall.Act("second", new() {
            Reference = reference, Action = HtmlAutomationAction.Click, WaitForReady = false
        }));

        Assert.True(first.IsSuccess);
        Assert.Equal(HtmlAutomationStatus.Success, first.Automation!.Status);
        Assert.True(second.IsSuccess);
        Assert.Equal(HtmlAutomationStatus.Stale, second.Automation!.Status);
    }

    [Theory]
    [InlineData("{\"css\":\"button\",\"action\":\"Inspect\",\"action\":\"Click\"}")]
    [InlineData("{\"css\":\"button\",\"action\":\"Click\",\"unexpected\":true}")]
    [InlineData("{\"css\":\"button\"}")]
    public async Task DispatcherRejectsAmbiguousOrIncompleteArgumentsBeforePageBehavior(string argumentsJson) {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = "<button>Increment</button><output>0</output>",
            Scripts = new[] { "document.querySelector('button').onclick=()=>document.querySelector('output').textContent='1'" }
        });
        using JsonDocument arguments = JsonDocument.Parse(argumentsJson);
        var call = new HtmlAutomationToolCall("untrusted", HtmlAutomationToolNames.Act, arguments.RootElement);

        HtmlAutomationToolResult result = await new HtmlAutomationToolDispatcher().ExecuteAsync(page, call);

        Assert.False(result.IsSuccess);
        Assert.NotNull(result.Error);
        Assert.Equal("0", (await page.EvaluateAsync("document.querySelector('output').textContent")).GetString());
    }

    [Fact]
    public async Task ActionToolPreservesKeyboardModifiers() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = "<input id='target'><output></output>",
            Scripts = new[] { "document.querySelector('#target').onkeydown=e=>document.querySelector('output').textContent=String(e.ctrlKey)+'/'+String(e.shiftKey)+'/'+e.key" }
        });
        var dispatcher = new HtmlAutomationToolDispatcher();

        HtmlAutomationToolResult result = await dispatcher.ExecuteAsync(page, HtmlAutomationToolCall.Act("press", new() {
            Query = HtmlLocatorQuery.Css("#target"), Action = HtmlAutomationAction.Press,
            Value = "K", Modifiers = HtmlKeyboardModifiers.Control | HtmlKeyboardModifiers.Shift,
            WaitForReady = false
        }));

        Assert.True(result.IsSuccess);
        Assert.Equal(HtmlAutomationStatus.Success, result.Automation!.Status);
        Assert.Equal("true/true/K", (await page.EvaluateAsync("document.querySelector('output').textContent")).GetString());
        JsonElement schema = HtmlAutomationToolCatalog.GetDefinitions().Single(tool => tool.Name == HtmlAutomationToolNames.Act).InputSchema;
        Assert.True(schema.GetProperty("properties").TryGetProperty("modifiers", out _));
    }

    [Fact]
    public async Task PublicRuntimeResultsSerializeWithoutReflectionMetadata() {
        HtmlProcessRuntimeProvider host = Runtime();
        await using IHtmlRuntimeContext context = await host.CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() { Html = "<p>Serializable</p>" });
        HtmlPageObservation observation = await page.ObserveAsync();
        HtmlAutomationResult automation = await page.AutomateAsync(new() {
            Query = HtmlLocatorQuery.Css("p"), Action = HtmlAutomationAction.Inspect, WaitForReady = false
        });
        HtmlScriptCapture capture = await page.CaptureAsync();
        var tool = new HtmlAutomationToolResult {
            CallId = "capture", ToolName = HtmlAutomationToolNames.Capture, IsSuccess = true, Capture = capture
        };
        var run = new HtmlAutomationRunResult {
            IsComplete = true, Steps = 1, FinalObservation = observation, ToolResults = new[] { tool }
        };

        string[] outputs = { HtmlRuntimeJson.Serialize(host.Descriptor), HtmlRuntimeJson.Serialize(observation),
            HtmlRuntimeJson.Serialize(automation), HtmlRuntimeJson.Serialize(page.GetTrace()),
            HtmlRuntimeJson.Serialize(capture), HtmlRuntimeJson.Serialize(tool), HtmlRuntimeJson.Serialize(run) };

        Assert.All(outputs, json => Assert.Equal(JsonValueKind.Object, JsonDocument.Parse(json).RootElement.ValueKind));
        Assert.Contains("Serializable", outputs[4], StringComparison.Ordinal);
    }

    [Fact]
    public async Task RulesDrivenPlannerCompletesThroughThePublicObservationActionLoop() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = "<button>Approve</button><output>Pending</output>",
            Scripts = new[] { "document.querySelector('button').onclick=()=>document.querySelector('output').textContent='Approved'" }
        });
        var runner = new HtmlAutomationRunner();

        HtmlAutomationRunResult result = await runner.RunAsync(page, (turn, _) => {
            if (turn.Observation.Elements.Any(element => element.Text == "Approved"))
                return Task.FromResult(HtmlAutomationPlannerDecision.Complete("approved"));
            HtmlObservedElement button = turn.Observation.Elements.Single(element => element.Role == "button" && element.AccessibleName == "Approve");
            return Task.FromResult(HtmlAutomationPlannerDecision.Execute(HtmlAutomationToolCall.Act("approve-" + turn.Step, new() {
                Reference = button.Reference,
                Action = HtmlAutomationAction.Click,
                WaitForReady = false
            })));
        }, new HtmlAutomationRunOptions {
            MaxSteps = 3,
            Observation = new HtmlPageObservationRequest { Mode = HtmlPageObservationMode.Combined, IncludeHidden = true }
        });

        Assert.True(result.IsComplete);
        Assert.Equal("approved", result.Message);
        Assert.Equal(2, result.Steps);
        Assert.Single(result.ToolResults);
        Assert.Equal(HtmlAutomationStatus.Success, result.ToolResults[0].Automation!.Status);
    }

    [Fact]
    public async Task PlannerStopsAtItsBoundAndRejectsEmptyContinuations() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() { Html = "<p>Idle</p>" });
        var runner = new HtmlAutomationRunner();

        await Assert.ThrowsAsync<InvalidOperationException>(() => runner.RunAsync(page,
            (_, _) => Task.FromResult(new HtmlAutomationPlannerDecision()), new() { MaxSteps = 1 }));

        HtmlAutomationRunResult bounded = await runner.RunAsync(page,
            (turn, _) => Task.FromResult(HtmlAutomationPlannerDecision.Execute(HtmlAutomationToolCall.Observe("observe-" + turn.Step))),
            new() { MaxSteps = 2 });
        Assert.False(bounded.IsComplete);
        Assert.Equal(2, bounded.Steps);
        Assert.Equal(2, bounded.ToolResults.Count);
    }
}
