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
    /// <summary>Runs navigation, observation, locator, action, wait, extraction and capture through public contracts.</summary>
    public static async Task<HtmlRuntimeConformanceReport> RunAsync(IHtmlRuntimeHost host, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(host);
        var cases = new List<HtmlRuntimeConformanceCaseResult> {
            await RunCaseAsync("observation-action-wait-extraction-capture", token => InteractionAsync(host, token), cancellationToken).ConfigureAwait(false),
            await RunCaseAsync("navigation", token => NavigationAsync(host, token), cancellationToken).ConfigureAwait(false),
            await RunCaseAsync("stale-reference", token => StaleReferenceAsync(host, token), cancellationToken).ConfigureAwait(false)
        };
        return new HtmlRuntimeConformanceReport { Provider = host.Descriptor, Cases = Array.AsReadOnly(cases.ToArray()) };
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
    }

    private static async Task<HtmlRuntimeConformanceCaseResult> RunCaseAsync(string id, Func<CancellationToken, Task> action, CancellationToken token) {
        var timer = Stopwatch.StartNew();
        try {
            await action(token).ConfigureAwait(false);
            return new HtmlRuntimeConformanceCaseResult { Id = id, Passed = true, Elapsed = timer.Elapsed };
        } catch (Exception error) when (error is not OperationCanceledException) {
            return new HtmlRuntimeConformanceCaseResult { Id = id, Error = error.Message, Elapsed = timer.Elapsed };
        }
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new HtmlRuntimeConformanceException(message);
    }
}

/// <summary>A provider failed a black-box conformance assertion.</summary>
public sealed class HtmlRuntimeConformanceException : InvalidOperationException {
    /// <summary>Creates a conformance assertion failure.</summary>
    public HtmlRuntimeConformanceException(string message) : base(message) { }
}
