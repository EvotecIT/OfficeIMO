using global::OfficeIMO.AI;
using global::OfficeIMO.AI.Html;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Conformance;

if (args.Length != 1 || !File.Exists(args[0]))
    throw new ArgumentException("Pass the deployed OfficeIMO.Html.Runtime.Worker.dll path.");

var host = new HtmlProcessRuntimeProvider(args[0], new UnusedDomServices());
if (!host.Descriptor.Supports(HtmlRuntimeCapabilityIds.RevisionBoundReferences))
    throw new InvalidOperationException("The NativeAOT client lost runtime capabilities.");
if (HtmlAutomationToolCatalog.GetDefinitions().Count != 4)
    throw new InvalidOperationException("The NativeAOT client lost automation tool schemas.");
var aiPlanner = new HtmlAutomationAiPlanner(new UnusedAiExecutor(), _ => "NativeAOT smoke");
if (aiPlanner.Tools.Count != 4)
    throw new InvalidOperationException("The NativeAOT client lost OfficeIMO AI tool declarations.");

await using IHtmlRuntimeContext context = await host.CreateContextAsync(new HtmlRuntimeContextOptions { Id = "native-aot" });
await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
    Profile = HtmlRuntimeProfile.WebApplicationV1,
    Html = "<button>Run</button><output>Pending</output>",
    Scripts = new[] { "document.querySelector('button').onclick=()=>document.querySelector('output').textContent='Complete'" }
});
HtmlPageObservation observation = await page.ObserveAsync(new HtmlPageObservationRequest { ActionableOnly = true });
HtmlObservedElement button = observation.Elements.Single();
HtmlAutomationResult action = await page.AutomateAsync(new HtmlAutomationRequest {
    Reference = button.Reference,
    Action = HtmlAutomationAction.Click,
    WaitForReady = false
});
action.EnsureSuccess();
var capture = new HtmlScriptCapture(new HtmlDocument(new UnusedDomServices(), "aot-capture").Freeze(), "aot-smoke",
    resources: new[] { HtmlRuntimeResource.FromText(new Uri("https://officeimo.invalid/data.json"), "{}", "application/json") });
var toolResult = new HtmlAutomationToolResult { CallId = "aot", ToolName = HtmlAutomationToolNames.Capture, IsSuccess = true, Capture = capture };
var runResult = new HtmlAutomationRunResult { IsComplete = true, Steps = 1, FinalObservation = observation, ToolResults = new[] { toolResult } };
HtmlRuntimeQualificationManifest automationManifest = HtmlRuntimeQualificationCatalog.Get("programmatic-automation-v1");
HtmlRuntimeConsumerQualificationResult consumerEvidence = HtmlRuntimeQualificationCatalog.EvaluateConsumer(
    automationManifest, "officeimo-ai-html", passedWorkflows: 1, failedWorkflows: 0);
var qualificationReport = new HtmlRuntimeConformanceReport {
    Provider = host.Descriptor,
    Cases = automationManifest.Cases.Select(item => new HtmlRuntimeConformanceCaseResult {
        Id = item.Id, Passed = true, RequiredAssertions = item.RequiredAssertions, PassedAssertions = item.RequiredAssertions
    }).ToArray()
};
HtmlRuntimeProfileQualificationResult profileEvidence = HtmlRuntimeQualificationCatalog.Evaluate(
    automationManifest, qualificationReport, consumerEvidence);
if (!profileEvidence.Passed)
    throw new InvalidOperationException("The NativeAOT client lost combined profile qualification.");
string[] jsonOutputs = {
    HtmlRuntimeJson.Serialize(host.Descriptor),
    HtmlRuntimeJson.Serialize(observation),
    HtmlRuntimeJson.Serialize(action),
    HtmlRuntimeJson.Serialize(page.GetTrace()),
    HtmlRuntimeJson.Serialize(capture.ArtifactManifest),
    HtmlRuntimeJson.Serialize(capture),
    HtmlRuntimeJson.Serialize(toolResult),
    HtmlRuntimeJson.Serialize(runResult),
    HtmlRuntimeConformanceJson.Serialize(new HtmlRuntimeConformanceReport {
        Provider = host.Descriptor,
        Cases = new[] { new HtmlRuntimeConformanceCaseResult { Id = "aot-json", Passed = true } }
    }),
    HtmlRuntimeConformanceJson.Serialize(HtmlRuntimeQualificationCatalog.Get("web-application-v1")),
    HtmlRuntimeConformanceJson.Serialize(consumerEvidence),
    HtmlRuntimeConformanceJson.Serialize(profileEvidence)
};
if (jsonOutputs.Any(json => string.IsNullOrWhiteSpace(json) || json[0] != '{'))
    throw new InvalidOperationException("The NativeAOT client lost public runtime JSON serialization.");
if ((await page.EvaluateAsync("document.querySelector('output').textContent")).GetString() != "Complete")
    throw new InvalidOperationException("The NativeAOT client lost the runtime observation/action protocol.");

file sealed class UnusedDomServices : IHtmlDomServices {
    public string Serialize(HtmlNode node, bool childrenOnly = false) => node.TextContent;
    public IReadOnlyList<HtmlElement> QuerySelectorAll(HtmlNode scope, string selector) => throw new NotSupportedException();
    public bool Matches(HtmlElement element, string selector) => throw new NotSupportedException();
}

file sealed class UnusedAiExecutor : IOfficeAiExecutor {
    public OfficeAiExecutionProfile Profile { get; } = new() {
        Id = "native-aot",
        Provider = "test",
        Model = "test",
        IsLocal = true,
        EnforcesJsonSchema = true
    };

    public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request,
        CancellationToken cancellationToken = default) => throw new NotSupportedException();
}
