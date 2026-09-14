using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeHostObservationTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task ProviderDescribesCapabilitiesBeforeStartingAContextAndEnforcesItsPageLimit() {
        var host = Runtime();

        Assert.Equal("officeimo.trusted-process", host.Descriptor.Id);
        Assert.Equal(1, host.Descriptor.MaximumPagesPerContext);
        Assert.True(host.Descriptor.Supports(HtmlRuntimeCapabilityIds.SemanticObservation));
        Assert.True(host.Descriptor.Supports(HtmlRuntimeCapabilityIds.RevisionBoundReferences));
        Assert.False(host.Descriptor.Supports(HtmlRuntimeCapabilityIds.ScreenshotObservation));
        Assert.Equal(typeof(HtmlRuntimeProviderDescriptor).Assembly.GetName().Version?.ToString(), host.Descriptor.Version);

        await using IHtmlRuntimeContext context = await host.CreateContextAsync(new() { Id = "qualification" });
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() { Html = "<p>First</p>" });

        Assert.Equal("qualification", page.ContextId);
        Assert.Single(context.Pages);
        await Assert.ThrowsAsync<InvalidOperationException>(() => context.OpenPageAsync(new() { Html = "<p>Second</p>" }));
    }

    [Fact]
    public void ProviderRejectsAWorkerWithoutItsCompatibilityManifest() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-runtime-manifest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string worker = Path.Combine(directory, "OfficeIMO.Html.Runtime.Worker.dll");
            File.Copy(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), worker);
            Assert.Throws<FileNotFoundException>(() => new HtmlProcessRuntimeProvider(worker, AngleSharpDomServices.Instance));
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public void ProviderRejectsManifestCapabilitiesThatDoNotMatchTheWorkerProtocol() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-runtime-capabilities-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string source = Path.Combine(AppContext.BaseDirectory, "RuntimeWorker");
            string worker = Path.Combine(directory, "OfficeIMO.Html.Runtime.Worker.dll");
            File.Copy(Path.Combine(source, "OfficeIMO.Html.Runtime.Worker.dll"), worker);
            string manifest = File.ReadAllText(Path.Combine(source, "OfficeIMO.Html.Runtime.Worker.manifest.json"))
                .Replace("trace.operations", "observation.screenshot", StringComparison.Ordinal);
            File.WriteAllText(Path.Combine(directory, "OfficeIMO.Html.Runtime.Worker.manifest.json"), manifest);

            Assert.Throws<NotSupportedException>(() => new HtmlProcessRuntimeProvider(worker, AngleSharpDomServices.Instance));
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task CombinedObservationReturnsBoundedSemanticAndVisualState() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri("https://runtime.officeimo.test/report"),
            Html = "<!doctype html><title>Report</title><main><h1>Quarterly report</h1><button id='approve'>Approve</button><input aria-label='Owner' value='Ada'></main>"
        });

        HtmlPageObservation observation = await page.ObserveAsync(new() {
            Mode = HtmlPageObservationMode.Combined,
            MaxElements = 32,
            MaxTextCharacters = 1024
        });

        Assert.Equal(page.Id, observation.PageId);
        Assert.Equal(page.ContextId, observation.ContextId);
        Assert.Equal("Report", observation.Title);
        Assert.Equal(new Uri("https://runtime.officeimo.test/report"), observation.Url);
        Assert.True(observation.Revision > 0);
        Assert.False(observation.IsTruncated);
        HtmlObservedElement button = Assert.Single(observation.Elements, item => item.Reference.ElementId == "approve");
        Assert.Equal("button", button.Role);
        Assert.Equal("Approve", button.AccessibleName);
        Assert.True(button.IsActionable);
        Assert.True(button.IsVisible);
        Assert.NotNull(button.BoundingBox);
        Assert.Contains(observation.Elements, item => item.AccessibleName == "Owner" && item.Value == "Ada");
    }

    [Fact]
    public async Task ObservedReferenceActsOnceAndBecomesStaleAfterThePageChanges() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = "<button id='go'>Go</button><output>0</output>",
            Scripts = new[] { "document.querySelector('button').onclick=()=>document.querySelector('output').textContent='1'" }
        });
        HtmlObservedElementReference reference = (await page.ObserveAsync(new() { ActionableOnly = true })).Elements.Single().Reference;

        HtmlAutomationResult clicked = await page.AutomateAsync(new() {
            Reference = reference,
            Action = HtmlAutomationAction.Click,
            WaitForReady = false
        });
        HtmlAutomationResult stale = await page.AutomateAsync(new() {
            Reference = reference,
            Action = HtmlAutomationAction.Inspect,
            WaitForReady = false
        });

        Assert.Equal(HtmlAutomationStatus.Success, clicked.Status);
        Assert.True(clicked.PageRevision > reference.Revision);
        Assert.Equal(HtmlAutomationStatus.Stale, stale.Status);
        Assert.Equal("1", (await page.Locator("output").InspectAsync()).Text);
    }

    [Fact]
    public async Task ObservationModesAndBoundsDoNotLeakUnrequestedData() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = "<button>Secret label</button><p>Additional text</p>"
        });

        HtmlPageObservation visual = await page.ObserveAsync(new() {
            Mode = HtmlPageObservationMode.Visual,
            MaxElements = 1,
            MaxTextCharacters = 4,
            IncludeHidden = true
        });

        Assert.True(visual.IsTruncated);
        Assert.Single(visual.Elements);
        Assert.All(visual.Elements, element => {
            Assert.Empty(element.Text);
            Assert.Empty(element.AccessibleName);
            Assert.NotNull(element.IsVisible);
        });
        await Assert.ThrowsAsync<NotSupportedException>(() => page.ObserveAsync(new() { IncludeScreenshotReference = true }));
    }

    [Fact]
    public async Task PasswordValuesAndSelectionAreOmittedFromObservationsAndActionResults() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync();
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = "<input id='secret' type='password' value='correct horse battery staple' aria-label='Password'>"
        });

        HtmlObservedElement observed = (await page.ObserveAsync(new() { ActionableOnly = true })).Elements.Single();
        HtmlAutomationResult inspected = await page.AutomateAsync(new() {
            Query = HtmlLocatorQuery.Css("#secret"), Action = HtmlAutomationAction.Inspect, WaitForReady = false
        });
        HtmlAutomationResult filled = await page.AutomateAsync(new() {
            Query = HtmlLocatorQuery.Css("#secret"), Action = HtmlAutomationAction.Fill,
            Value = "new secret", WaitForReady = false
        });

        Assert.Null(observed.Value);
        Assert.Null(observed.SelectionStart);
        Assert.Null(observed.SelectionEnd);
        Assert.Null(inspected.Element!.Value);
        Assert.Null(filled.Element!.Value);
        Assert.Equal("new secret", (await page.EvaluateAsync("document.querySelector('#secret').value")).GetString());
    }

    [Fact]
    public async Task TraceIsBoundedAndRedactsOptionalUrlDetails() {
        var options = new HtmlRuntimeContextOptions {
            Trace = new HtmlRuntimeTraceOptions {
                MaxEvents = 8,
                IncludeUrls = true,
                Redactor = value => value.Replace("secret", "redacted", StringComparison.Ordinal)
            }
        };
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync(options);
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri("https://runtime.officeimo.test/start"),
            Html = "<p>Ready</p>"
        });
        await page.NavigateAsync(new Uri("https://runtime.officeimo.test/start#secret"));
        for (int i = 0; i < 10; i++) await page.ObserveAsync();

        HtmlRuntimeTrace trace = page.GetTrace();
        Assert.True(trace.IsTruncated);
        Assert.Equal(8, trace.Events.Count);
        Assert.Equal(Enumerable.Range(1, 8).Select(value => (long)value), trace.Events.Select(item => item.Sequence));
        Assert.DoesNotContain(trace.Events, item => item.Detail?.Contains("secret", StringComparison.Ordinal) == true);
        Assert.Contains(trace.Events, item => item.Detail?.Contains("redacted", StringComparison.Ordinal) == true);
        Assert.DoesNotContain(trace.Events, item => item.Url?.AbsoluteUri.Contains("secret", StringComparison.Ordinal) == true);
    }
}
