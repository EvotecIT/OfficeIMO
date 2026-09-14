using System.Collections.ObjectModel;
using System.Diagnostics;

namespace OfficeIMO.Html.Runtime.Conformance;

/// <summary>Outcome of one provider-neutral runtime conformance scenario.</summary>
public sealed class HtmlRuntimeConformanceCaseResult {
    /// <summary>Stable scenario identifier.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>Whether every assertion passed.</summary>
    public bool Passed { get; init; }
    /// <summary>Failure detail, or null for a passing case.</summary>
    public string? Error { get; init; }
    /// <summary>Elapsed scenario time.</summary>
    public TimeSpan Elapsed { get; init; }
    /// <summary>Exact assertions required by this versioned scenario.</summary>
    public int RequiredAssertions { get; init; }
    /// <summary>Assertions reached and passed before completion or failure.</summary>
    public int PassedAssertions { get; init; }
    /// <summary>Assertions which failed.</summary>
    public int FailedAssertions { get; init; }
    /// <summary>Required assertions not reached after a failure.</summary>
    public int UntestedAssertions => Math.Max(0, RequiredAssertions - PassedAssertions - FailedAssertions);
}

/// <summary>Immutable conformance report for one advertised provider.</summary>
public sealed class HtmlRuntimeConformanceReport {
    /// <summary>Provider descriptor tested by the suite.</summary>
    public HtmlRuntimeProviderDescriptor Provider { get; init; } = null!;
    /// <summary>Scenario results.</summary>
    public IReadOnlyList<HtmlRuntimeConformanceCaseResult> Cases { get; init; } = Array.Empty<HtmlRuntimeConformanceCaseResult>();
    /// <summary>Whether every required scenario passed.</summary>
    public bool Passed => Cases.Count > 0 && Cases.All(result => result.Passed);
}

/// <summary>Reusable black-box scenarios for implementations of <see cref="IHtmlRuntimeHost"/>.</summary>
public static class HtmlRuntimeConformanceSuite {
    private static readonly AsyncLocal<AssertionCounter?> CurrentAssertions = new();
    private static readonly IReadOnlyList<string> OrderedCaseIds = Array.AsReadOnly(new[] {
        "scripted-resource-capture", "observation-action-wait-extraction-capture", "navigation",
        "stale-reference-recovery", "programmatic-rules-workflow", "provider-observability",
        "blocked-resource-script-failure"
    });
    private static readonly IReadOnlyDictionary<string, int> AssertionsByCase = new ReadOnlyDictionary<string, int>(new Dictionary<string, int>(StringComparer.Ordinal) {
        ["scripted-resource-capture"] = 4,
        ["observation-action-wait-extraction-capture"] = 13,
        ["navigation"] = 5,
        ["stale-reference-recovery"] = 3,
        ["programmatic-rules-workflow"] = 5,
        ["provider-observability"] = 8,
        ["blocked-resource-script-failure"] = 7
    });

    /// <summary>Stable case identifiers and exact assertion counts for this suite version.</summary>
    public static IReadOnlyDictionary<string, int> CaseAssertions => AssertionsByCase;

    /// <summary>Runs navigation, observation, locator, action, wait, extraction and capture through public contracts.</summary>
    public static Task<HtmlRuntimeConformanceReport> RunAsync(IHtmlRuntimeHost host, CancellationToken cancellationToken = default) =>
        RunSelectedAsync(host, OrderedCaseIds, cancellationToken);

    /// <summary>Runs exactly the cases selected by a versioned qualification manifest.</summary>
    public static Task<HtmlRuntimeConformanceReport> RunAsync(IHtmlRuntimeHost host,
        HtmlRuntimeQualificationManifest manifest, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(host);
        ArgumentNullException.ThrowIfNull(manifest);
        if (!manifest.Providers.Any(provider => provider.Id == host.Descriptor.Id))
            throw new InvalidOperationException($"Qualification profile '{manifest.Id}' has no expectation for provider '{host.Descriptor.Id}'.");
        return RunSelectedAsync(host, manifest.Cases.Select(item => item.Id), cancellationToken);
    }

    private static async Task<HtmlRuntimeConformanceReport> RunSelectedAsync(IHtmlRuntimeHost host,
        IEnumerable<string> caseIds, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(host);
        var cases = new List<HtmlRuntimeConformanceCaseResult>();
        foreach (string id in caseIds) {
            Func<CancellationToken, Task> action = id switch {
                "scripted-resource-capture" => token => ScriptedResourceCaptureAsync(host, token),
                "observation-action-wait-extraction-capture" => token => InteractionAsync(host, token),
                "navigation" => token => NavigationAsync(host, token),
                "stale-reference-recovery" => token => StaleReferenceAsync(host, token),
                "programmatic-rules-workflow" => token => RulesWorkflowAsync(host, token),
                "provider-observability" => token => ProviderObservabilityAsync(host, token),
                "blocked-resource-script-failure" => token => FailureDiagnosticsAsync(host, token),
                _ => throw new InvalidOperationException($"Unknown runtime conformance case '{id}'.")
            };
            cases.Add(await RunCaseAsync(id, action, cancellationToken).ConfigureAwait(false));
        }
        return new HtmlRuntimeConformanceReport { Provider = host.Descriptor, Cases = Array.AsReadOnly(cases.ToArray()) };
    }

    private static async Task ScriptedResourceCaptureAsync(IHtmlRuntimeHost host, CancellationToken token) {
        var data = new Uri("https://conformance.officeimo.test/data.json");
        HtmlRuntimeProfile profile = host.Descriptor.Profiles.Contains(HtmlRuntimeProfile.ScriptedDocumentV1)
            ? HtmlRuntimeProfile.ScriptedDocumentV1 : HtmlRuntimeProfile.WebApplicationV1;
        await using IHtmlRuntimeContext context = await host.CreateContextAsync(cancellationToken: token).ConfigureAwait(false);
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            Profile = profile,
            DocumentUrl = new Uri("https://conformance.officeimo.test/capture"),
            Html = "<!doctype html><title>Capture</title><output>Pending</output>",
            Resources = new[] { HtmlRuntimeResource.FromText(data, "{\"total\":42}", "application/json") },
            Scripts = new[] { "console.info('qualification marker');fetch('/data.json').then(r=>r.json()).then(v=>document.querySelector('output').textContent='Total '+v.total)" }
        }, token).ConfigureAwait(false);
        await page.WaitForAsync("document.querySelector('output').textContent === 'Total 42'", token).ConfigureAwait(false);
        HtmlScriptCapture first = await page.CaptureAsync(cancellationToken: token).ConfigureAwait(false);
        HtmlScriptCapture second = await page.CaptureAsync(cancellationToken: token).ConfigureAwait(false);
        Require(first.Document.QuerySelector("output")?.TextContent == "Total 42", "The resource-backed script did not produce the expected output.");
        Require(first.Resources.Any(item => item.Url == data && item.StatusCode == 200), "The capture did not retain the supplied resource response.");
        Require(first.ArtifactManifest.Entries.Count == 2 && first.ArtifactManifest.ByteCount > 0, "The deterministic manifest does not cover the document and resource.");
        Require(first.ArtifactManifest.Id == second.ArtifactManifest.Id, "Equivalent captures produced different artifact identities.");
    }

    private static async Task InteractionAsync(IHtmlRuntimeHost host, CancellationToken token) {
        await using IHtmlRuntimeContext context = await host.CreateContextAsync(cancellationToken: token).ConfigureAwait(false);
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri("https://conformance.officeimo.test/start"),
            Html = "<!doctype html><title>Conformance</title><button id='run'>Run</button><output>Pending</output><i id='temporary'>Temporary</i><input id='keys'><input id='password' type='password' value='must-not-leak'>",
            Scripts = new[] { "document.querySelector('#run').onclick=()=>{setTimeout(()=>document.querySelector('output').textContent='Complete',50);setTimeout(()=>document.querySelector('#temporary').remove(),500)};document.querySelector('#keys').addEventListener('keydown',e=>document.body.dataset.key=String(e.ctrlKey)+'/'+String(e.shiftKey)+'/'+e.key)" }
        }, token).ConfigureAwait(false);
        HtmlPageObservation observed = await page.ObserveAsync(new HtmlPageObservationRequest {
            Mode = HtmlPageObservationMode.Combined, IncludeHidden = true
        }, token).ConfigureAwait(false);
        Require(observed.ProviderId == host.Descriptor.Id, "Observation provider identity differs from the host descriptor.");
        HtmlObservedElement button = observed.Elements.Single(element => element.Reference.ElementId == "run");
        HtmlObservedElement password = observed.Elements.Single(element => element.Reference.ElementId == "password");
        Require(button.IsActionable && button.Role == "button", "The button is not exposed as actionable semantic content.");
        Require(password.Value == null && password.SelectionStart == null && password.SelectionEnd == null,
            "A password observation exposed its value or selection length.");
        HtmlAutomationResult action = await page.AutomateAsync(new HtmlAutomationRequest {
            Reference = button.Reference, Action = HtmlAutomationAction.Click, WaitForReady = false
        }, token).ConfigureAwait(false);
        Require(action.Status == HtmlAutomationStatus.Success, "The reference-bound click failed: " + action.Status);
        HtmlAutomationResult wait = await page.AutomateAsync(new HtmlAutomationRequest {
            Query = HtmlLocatorQuery.Css("output"), Action = HtmlAutomationAction.Wait,
            WaitState = HtmlLocatorWaitState.Text, Value = "Complete"
        }, token).ConfigureAwait(false);
        Require(wait.Status == HtmlAutomationStatus.Success, "The structured text wait failed.");
        HtmlAutomationResult stillAttached = await page.AutomateAsync(new HtmlAutomationRequest {
            Query = HtmlLocatorQuery.Css("#temporary"), Action = HtmlAutomationAction.Inspect, WaitForReady = false
        }, token).ConfigureAwait(false);
        Require(stillAttached.Status == HtmlAutomationStatus.Success, "The detached-wait target was not attached before polling began.");
        HtmlAutomationResult detached = await page.AutomateAsync(new HtmlAutomationRequest {
            Query = HtmlLocatorQuery.Css("#temporary"), Action = HtmlAutomationAction.Wait,
            WaitState = HtmlLocatorWaitState.Detached
        }, token).ConfigureAwait(false);
        Require(detached.Status == HtmlAutomationStatus.Success, "The delayed detached wait failed.");
        HtmlAutomationResult missingHidden = await page.AutomateAsync(new HtmlAutomationRequest {
            Query = HtmlLocatorQuery.Css("#never-present"), Action = HtmlAutomationAction.Wait,
            WaitState = HtmlLocatorWaitState.Hidden
        }, token).ConfigureAwait(false);
        Require(missingHidden.Status == HtmlAutomationStatus.Success, "A missing target did not satisfy the hidden wait.");
        HtmlAutomationResult modifiedPress = await page.AutomateAsync(new HtmlAutomationRequest {
            Query = HtmlLocatorQuery.Css("#keys"), Action = HtmlAutomationAction.Press, Value = "K",
            Modifiers = HtmlKeyboardModifiers.Control | HtmlKeyboardModifiers.Shift, WaitForReady = false
        }, token).ConfigureAwait(false);
        Require(modifiedPress.Status == HtmlAutomationStatus.Success, "The keyboard-modifier action failed.");
        Require((await page.EvaluateAsync("document.body.dataset.key", token).ConfigureAwait(false)).GetString() == "true/true/K",
            "The provider did not preserve keyboard modifiers.");
        Require((await page.EvaluateAsync("document.querySelector('output').textContent", token).ConfigureAwait(false)).GetString() == "Complete",
            "Expression extraction returned the wrong value.");
        HtmlScriptCapture capture = await page.CaptureAsync(cancellationToken: token).ConfigureAwait(false);
        Require(capture.Document.QuerySelector("output")?.TextContent == "Complete", "The inert capture lost the completed DOM state.");
        await page.WaitForAsync("true", token).ConfigureAwait(false);
        HtmlRuntimeTrace trace = page.GetTrace();
        Require(new[] { HtmlRuntimeEventKind.Observation, HtmlRuntimeEventKind.Action, HtmlRuntimeEventKind.Script,
                HtmlRuntimeEventKind.Wait, HtmlRuntimeEventKind.Capture }
            .All(kind => trace.Events.Any(item => item.Kind == kind && item.Status == "success")),
            "The operation trace does not cover observation, action, script, wait and capture calls.");
    }

    private static async Task NavigationAsync(IHtmlRuntimeHost host, CancellationToken token) {
        var next = new Uri("https://conformance.officeimo.test/next");
        var replacement = new Uri("https://conformance.officeimo.test/replacement");
        await using IHtmlRuntimeContext context = await host.CreateContextAsync(cancellationToken: token).ConfigureAwait(false);
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri("https://conformance.officeimo.test/start"),
            Html = "<!doctype html><a href='/next'>Next</a>",
            Resources = new[] {
                HtmlRuntimeResource.FromText(next, "<!doctype html><title>Next</title><h1>Destination</h1>", "text/html; charset=utf-8"),
                HtmlRuntimeResource.FromText(replacement, "<!doctype html><title>Replacement</title><h1>Replacement</h1>", "text/html; charset=utf-8")
            }
        }, token).ConfigureAwait(false);
        int initialHistoryLength = (await page.EvaluateAsync("history.length", token).ConfigureAwait(false)).GetInt32();
        await page.NavigateAsync(next, cancellationToken: token).ConfigureAwait(false);
        HtmlPageObservation observed = await page.ObserveAsync(new HtmlPageObservationRequest { Mode = HtmlPageObservationMode.Semantic, IncludeHidden = true }, token).ConfigureAwait(false);
        Require(observed.Url == next && observed.Title == "Next", "Navigation did not replace the page identity and content.");
        Require(observed.Elements.Any(element => element.Role == "heading" && element.Text == "Destination"), "Destination content was not observed.");
        await page.NavigateAsync(replacement, replaceHistoryEntry: true, cancellationToken: token).ConfigureAwait(false);
        HtmlPageObservation replaced = await page.ObserveAsync(new HtmlPageObservationRequest { Mode = HtmlPageObservationMode.Semantic, IncludeHidden = true }, token).ConfigureAwait(false);
        Require(replaced.Url == replacement && replaced.Title == "Replacement", "Replacement navigation did not load the requested document.");
        int replacedHistoryLength = (await page.EvaluateAsync("history.length", token).ConfigureAwait(false)).GetInt32();
        Require(replacedHistoryLength == initialHistoryLength + 1, "Replacement navigation appended a history entry.");
        await page.ReloadAsync(token).ConfigureAwait(false);
        HtmlRuntimeTrace trace = page.GetTrace();
        Require(trace.Events.Count(item => item.Kind == HtmlRuntimeEventKind.Navigation
                && item.Operation.StartsWith("navigate", StringComparison.Ordinal) && item.Status == "success") >= 2
            && trace.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Navigation && item.Operation == "reload" && item.Status == "success"),
            "The operation trace does not cover navigation, replacement and reload calls.");
    }

    private static async Task StaleReferenceAsync(IHtmlRuntimeHost host, CancellationToken token) {
        await using IHtmlRuntimeContext context = await host.CreateContextAsync(cancellationToken: token).ConfigureAwait(false);
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1, Html = "<button id='replace'>Replace</button>"
        }, token).ConfigureAwait(false);
        HtmlObservedElementReference reference = (await page.ObserveAsync(new HtmlPageObservationRequest { ActionableOnly = true }, token).ConfigureAwait(false)).Elements.Single().Reference;
        await page.ExecuteAsync("document.querySelector('#replace').outerHTML='<button id=\"new\">New</button>'", token).ConfigureAwait(false);
        HtmlAutomationResult result = await page.AutomateAsync(new HtmlAutomationRequest {
            Reference = reference, Action = HtmlAutomationAction.Inspect, WaitForReady = false
        }, token).ConfigureAwait(false);
        Require(result.Status == HtmlAutomationStatus.Stale, "A reference from an earlier page revision was accepted.");
        HtmlObservedElementReference current = (await page.ObserveAsync(new HtmlPageObservationRequest { ActionableOnly = true }, token).ConfigureAwait(false)).Elements.Single().Reference;
        Require(current.ElementId == "new" && current.Revision > reference.Revision, "A fresh observation did not expose the replacement revision.");
        HtmlAutomationResult recovered = await page.AutomateAsync(new HtmlAutomationRequest { Reference = current, Action = HtmlAutomationAction.Inspect, WaitForReady = false }, token).ConfigureAwait(false);
        Require(recovered.Status == HtmlAutomationStatus.Success, "The workflow did not recover with a fresh reference.");
    }

    private static async Task RulesWorkflowAsync(IHtmlRuntimeHost host, CancellationToken token) {
        await using IHtmlRuntimeContext context = await host.CreateContextAsync(cancellationToken: token).ConfigureAwait(false);
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri("https://conformance.officeimo.test/rules"),
            Html = "<!doctype html><label>Name <input id='name'></label><button id='submit'>Submit</button><output>Pending</output>",
            Scripts = new[] { "document.querySelector('#submit').onclick=()=>setTimeout(()=>document.querySelector('output').textContent='Accepted '+document.querySelector('#name').value,25)" }
        }, token).ConfigureAwait(false);
        var runner = new HtmlAutomationRunner();
        HtmlAutomationRunResult result = await runner.RunAsync(page, (turn, _) => Task.FromResult(turn.Step switch {
            0 => HtmlAutomationPlannerDecision.Execute(HtmlAutomationToolCall.Act("fill", new HtmlAutomationRequest { Query = HtmlLocatorQuery.Css("#name"), Action = HtmlAutomationAction.Fill, Value = "Ada" })),
            1 => HtmlAutomationPlannerDecision.Execute(HtmlAutomationToolCall.Act("submit", new HtmlAutomationRequest { Query = HtmlLocatorQuery.Css("#submit"), Action = HtmlAutomationAction.Click, WaitForReady = false })),
            2 => HtmlAutomationPlannerDecision.Execute(
                HtmlAutomationToolCall.Act("wait", new HtmlAutomationRequest { Query = HtmlLocatorQuery.Css("output"), Action = HtmlAutomationAction.Wait, WaitState = HtmlLocatorWaitState.Text, Value = "Accepted Ada" }),
                HtmlAutomationToolCall.Capture("capture")),
            _ => HtmlAutomationPlannerDecision.Complete("accepted")
        }), new HtmlAutomationRunOptions { MaxSteps = 5, MaxCallsPerStep = 2 }, token).ConfigureAwait(false);
        Require(result.IsComplete && result.Message == "accepted", "The deterministic planner did not complete its goal.");
        Require(result.Steps == 4, "The deterministic planner used an unexpected number of decisions.");
        Require(result.ToolResults.Count == 4 && result.ToolResults.All(item => item.IsSuccess), "The deterministic planner did not execute the expected successful tool sequence.");
        HtmlScriptCapture capture = result.ToolResults.Single(item => item.CallId == "capture").Capture!;
        Require(capture.Document.QuerySelector("output")?.TextContent == "Accepted Ada", "The planner capture lost the completed form result.");
        Require(capture.ArtifactManifest.Id.StartsWith("sha256:", StringComparison.Ordinal), "The planner capture has no content-addressed artifact identity.");
    }

    private static async Task ProviderObservabilityAsync(IHtmlRuntimeHost host, CancellationToken token) {
        var source = new Uri("https://conformance.officeimo.test/diagnostics-start.js?token=secret");
        var script = new Uri("https://conformance.officeimo.test/diagnostics.js?token=secret");
        await using IHtmlRuntimeContext context = await host.CreateContextAsync(new HtmlRuntimeContextOptions {
            Trace = new HtmlRuntimeTraceOptions { MaxEvents = 64, IncludeUrls = true, IncludeConsoleMessages = true,
                Redactor = value => value.Replace("secret", "redacted", StringComparison.Ordinal) }
        }, token).ConfigureAwait(false);
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri("https://conformance.officeimo.test/observability"),
            Html = "<!doctype html><body><a id='download' download='report.txt' href='data:text/plain,report'>Download</a><script src='/diagnostics-start.js?token=secret'></script></body>",
            Resources = new[] {
                new HtmlRuntimeResource(source, Array.Empty<byte>(), "text/javascript", 302,
                    headers: new Dictionary<string, string> { ["Location"] = "/diagnostics.js?token=secret" }),
                HtmlRuntimeResource.FromText(script, "console.info('secret marker')", "text/javascript")
            }
        }, token).ConfigureAwait(false);
        await page.ExecuteAsync("console.info('secret marker')", token).ConfigureAwait(false);
        HtmlAutomationResult download = await page.AutomateAsync(new HtmlAutomationRequest {
            Query = HtmlLocatorQuery.Css("#download"), Action = HtmlAutomationAction.Click, WaitForReady = false
        }, token).ConfigureAwait(false);
        Require(download.Status is HtmlAutomationStatus.Success or HtmlAutomationStatus.Unsupported,
            "The provider returned an unexpected bounded-download action status.");
        HtmlScriptCapture capture = await page.CaptureAsync(cancellationToken: token).ConfigureAwait(false);
        HtmlRuntimeTrace trace = page.GetTrace();
        Require(trace.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Console && item.Detail == "redacted marker"), "The trace has no redacted console event.");
        Require(trace.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Resource && item.StatusCode == 200), "The trace has no completed resource event.");
        Require(trace.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Redirect && item.StatusCode == 302), "The trace has no followed redirect event.");
        Require(trace.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Policy), "The trace has no resource policy decision.");
        Require(trace.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Download), "The trace has no bounded download event.");
        Require(trace.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Artifact && item.ArtifactId == capture.ArtifactManifest.Id), "The trace has no deterministic capture artifact event.");
        Require(!HtmlRuntimeJson.Serialize(trace).Contains("secret", StringComparison.Ordinal), "Trace redaction retained a sensitive marker.");
    }

    private static async Task FailureDiagnosticsAsync(IHtmlRuntimeHost host, CancellationToken token) {
        HtmlRuntimeProfile profile = host.Descriptor.Profiles.Contains(HtmlRuntimeProfile.ScriptedDocumentV1)
            ? HtmlRuntimeProfile.ScriptedDocumentV1 : HtmlRuntimeProfile.WebApplicationV1;
        HtmlRuntimeContextOptions traceOptions = new() { Trace = new HtmlRuntimeTraceOptions {
            MaxEvents = 64, IncludeConsoleMessages = true, IncludeFailureMessages = true,
            Redactor = value => value.Replace("secret", "redacted", StringComparison.Ordinal)
        } };
        await using (IHtmlRuntimeContext blockedContext = await host.CreateContextAsync(traceOptions, token).ConfigureAwait(false)) {
            await using IHtmlRuntimePage blockedPage = await blockedContext.OpenPageAsync(new HtmlScriptRequest {
                Profile = profile,
                DocumentUrl = new Uri("https://conformance.officeimo.test/blocked"),
                Html = "<!doctype html><body></body>",
                Scripts = new[] { "fetch('https://blocked.officeimo.test/private?token=secret').catch(()=>{console.warn('secret blocked');document.body.dataset.done='yes'})" }
            }, token).ConfigureAwait(false);
            await blockedPage.WaitForAsync("document.body.dataset.done === 'yes'", token).ConfigureAwait(false);
            HtmlRuntimeTrace blocked = blockedPage.GetTrace();
            Require(blocked.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Console && item.Detail == "redacted blocked"), "The blocked request did not produce a redacted console diagnostic.");
            Require(blocked.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Policy && item.Status == "blocked"), "The trace has no blocked resource policy decision.");
            Require(blocked.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Resource && item.Status == "failure"), "The trace has no failed resource event.");
        }
        await using IHtmlRuntimeContext failureContext = await host.CreateContextAsync(traceOptions, token).ConfigureAwait(false);
        await using IHtmlRuntimePage failurePage = await failureContext.OpenPageAsync(new HtmlScriptRequest {
            Profile = profile, Html = "<!doctype html><p>Ready</p>"
        }, token).ConfigureAwait(false);
        Exception? failure = null;
        try { await failurePage.ExecuteAsync("console.error('secret console');throw new Error('secret script failure')", token).ConfigureAwait(false); }
        catch (Exception error) when (error is not OperationCanceledException) { failure = error; }
        Require(failure != null, "A failing script completed successfully.");
        HtmlRuntimeTrace failed = failurePage.GetTrace();
        Require(failed.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Console && item.Detail == "redacted console"), "The trace has no redacted error-console event.");
        Require(failed.Events.Any(item => item.Kind == HtmlRuntimeEventKind.Failure && item.Detail?.Contains("redacted", StringComparison.Ordinal) == true), "The trace has no redacted script-failure event.");
        Require(!HtmlRuntimeJson.Serialize(failed).Contains("secret", StringComparison.Ordinal), "Failure trace redaction retained a sensitive marker.");
    }

    private static async Task<HtmlRuntimeConformanceCaseResult> RunCaseAsync(string id, Func<CancellationToken, Task> action, CancellationToken token) {
        var timer = Stopwatch.StartNew();
        var assertions = new AssertionCounter();
        CurrentAssertions.Value = assertions;
        try {
            await action(token).ConfigureAwait(false);
            int required = AssertionsByCase[id];
            if (assertions.Passed != required)
                throw new HtmlRuntimeConformanceException($"Scenario '{id}' executed {assertions.Passed} assertions; its manifest requires {required}.");
            return new HtmlRuntimeConformanceCaseResult { Id = id, Passed = true, Elapsed = timer.Elapsed,
                RequiredAssertions = required, PassedAssertions = assertions.Passed };
        } catch (Exception error) when (error is not OperationCanceledException) {
            return new HtmlRuntimeConformanceCaseResult { Id = id, Error = error.Message, Elapsed = timer.Elapsed,
                RequiredAssertions = AssertionsByCase[id], PassedAssertions = assertions.Passed, FailedAssertions = assertions.Failed };
        } finally { CurrentAssertions.Value = null; }
    }

    private static void Require(bool condition, string message) {
        AssertionCounter assertions = CurrentAssertions.Value ?? throw new InvalidOperationException("A conformance assertion ran outside a scenario.");
        if (condition) { assertions.Passed++; return; }
        assertions.Failed++;
        throw new HtmlRuntimeConformanceException(message);
    }

    private sealed class AssertionCounter { internal int Passed; internal int Failed; }
}

/// <summary>A provider failed a black-box conformance assertion.</summary>
public sealed class HtmlRuntimeConformanceException : InvalidOperationException {
    /// <summary>Creates a conformance assertion failure.</summary>
    public HtmlRuntimeConformanceException(string message) : base(message) { }
}
