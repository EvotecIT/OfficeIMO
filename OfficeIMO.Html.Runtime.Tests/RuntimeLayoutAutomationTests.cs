using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeLayoutAutomationTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    private static HtmlScriptRequest Application(string html) => new() {
        Profile = HtmlRuntimeProfile.WebApplicationV1,
        Html = html,
        DocumentUrl = new Uri("https://layout.example/"),
        ViewportWidth = 320,
        ViewportHeight = 160
    };

    [Fact]
    public async Task LayoutOnlyOperationsRejectTheSingleDocumentProfile() {
        await using var session = await Runtime().OpenTrustedAsync(new() { Html = "<button>Target</button>" });
        foreach (HtmlAutomationRequest request in new[] {
            new HtmlAutomationRequest { Query = HtmlLocatorQuery.Css("button"), Action = HtmlAutomationAction.ScrollIntoView, WaitForReady = false },
            new HtmlAutomationRequest { Query = HtmlLocatorQuery.Css("button"), Action = HtmlAutomationAction.Wait, WaitState = HtmlLocatorWaitState.Visible, WaitForReady = false },
            new HtmlAutomationRequest { Query = HtmlLocatorQuery.Css("button"), Action = HtmlAutomationAction.Press, Value = "Enter", WaitForReady = false },
            new HtmlAutomationRequest { Query = HtmlLocatorQuery.Css("button"), Action = HtmlAutomationAction.SetSelection, SelectionStart = 0, SelectionEnd = 0, WaitForReady = false }
        }) {
            Assert.Equal(HtmlAutomationStatus.Unsupported, (await session.AutomateAsync(request)).Status);
        }
    }

    [Fact]
    public async Task ComputedVisibilityAndPointerEventsGateTypedActions() {
        const string html = """
            <style>
              body { margin: 0 }
              button { display: block; width: 120px; height: 32px }
              #display-none { display: none }
              #visibility-hidden { visibility: hidden }
              #no-pointer { pointer-events: none }
            </style>
            <button id='visible'>Visible</button>
            <button id='display-none'>Display none</button>
            <button id='visibility-hidden'>Visibility hidden</button>
            <button id='no-pointer'>No pointer</button>
            """;
        await using var session = await Runtime().OpenTrustedAsync(Application(html));

        HtmlRuntimeElementState visible = await session.Locator("#visible").InspectAsync();
        Assert.True(visible.IsVisible);
        Assert.True(visible.IsInViewport);
        Assert.True(visible.AcceptsPointerEvents);
        Assert.NotNull(visible.BoundingBox);
        await session.Locator("#visible").WaitForVisibleAsync();

        foreach (string selector in new[] { "#display-none", "#visibility-hidden" }) {
            HtmlRuntimeElementState hidden = await session.Locator(selector).InspectAsync();
            Assert.False(hidden.IsVisible);
            Assert.Null(hidden.BoundingBox);
            await session.Locator(selector).WaitForHiddenAsync();
            var failure = await session.AutomateAsync(new() { Query = HtmlLocatorQuery.Css(selector), Action = HtmlAutomationAction.Click, WaitForReady = false });
            Assert.Equal(HtmlAutomationStatus.NotReady, failure.Status);
        }

        HtmlRuntimeElementState noPointer = await session.Locator("#no-pointer").InspectAsync();
        Assert.False(noPointer.AcceptsPointerEvents);
        var pointerFailure = await session.AutomateAsync(new() { Query = HtmlLocatorQuery.Css("#no-pointer"), Action = HtmlAutomationAction.Click, WaitForReady = false });
        Assert.Equal(HtmlAutomationStatus.NotReady, pointerFailure.Status);
    }

    [Fact]
    public async Task LocatorScrollAndActionUseTheOwnedLayoutViewport() {
        const string html = """
            <style>
              body { margin: 0 }
              #below { display: block; margin-top: 900px; width: 120px; height: 40px }
            </style>
            <button id='below' onclick="document.body.dataset.clicked='yes'">Below</button>
            """;
        await using var session = await Runtime().OpenTrustedAsync(Application(html));
        var below = session.Locator("#below");

        HtmlRuntimeElementState before = await below.InspectAsync();
        Assert.True(before.IsVisible);
        Assert.False(before.IsInViewport);
        Assert.NotNull(before.BoundingBox);
        Assert.True(before.BoundingBox!.Y >= 800D);

        await below.ScrollIntoViewAsync();
        HtmlRuntimeElementState scrolled = await below.InspectAsync();
        Assert.True(scrolled.ScrollY > 0D);
        Assert.True(scrolled.IsInViewport);
        Assert.True(scrolled.BoundingBox!.Y + scrolled.BoundingBox.Height <= 160.01D);
        await below.WaitForInViewportAsync();

        await session.ExecuteAsync("scrollTo(0,0)");
        await below.ClickAsync();
        Assert.Equal("yes", (await session.EvaluateAsync("document.body.dataset.clicked")).GetString());
        Assert.True((await below.InspectAsync()).ScrollY > 0D);
    }

    [Fact]
    public async Task ScriptGeometryAndScrollingShareTheAutomationViewport() {
        const string html = """
            <style>body{margin:0} #target{display:block;margin-top:700px;width:80px;height:30px}</style>
            <div id='target'>Target</div>
            """;
        await using var session = await Runtime().OpenTrustedAsync(Application(html));

        var result = await session.EvaluateAsync("(()=>{const target=document.querySelector('#target'),before=target.getBoundingClientRect();target.scrollIntoView();const after=target.getBoundingClientRect();return {before:before.top,after:after.top,width:after.width,scrollY,pageYOffset,innerHeight}})()");
        Assert.True(result.GetProperty("before").GetDouble() > 160D, result.ToString());
        Assert.InRange(result.GetProperty("after").GetDouble(), 0D, 160D);
        Assert.True(result.GetProperty("width").GetDouble() > 0D);
        Assert.True(result.GetProperty("scrollY").GetDouble() > 0D);
        Assert.Equal(result.GetProperty("scrollY").GetDouble(), result.GetProperty("pageYOffset").GetDouble());
        Assert.Equal(160D, result.GetProperty("innerHeight").GetDouble());
    }

    [Fact]
    public async Task LayoutInspectionRecomputesAfterStyleMutation() {
        await using var session = await Runtime().OpenTrustedAsync(Application(
            "<style>#target{display:none;width:80px;height:30px}</style><button id='target'>Target</button>"));
        var target = session.Locator("#target");
        await target.WaitForHiddenAsync();
        await session.ExecuteAsync("document.querySelector('#target').style.display='block'");
        await target.WaitForVisibleAsync();
        Assert.True((await target.InspectAsync()).IsVisible);
    }

    [Fact]
    public async Task LoadedExternalStylesheetParticipatesInOwnedActionabilityLayout() {
        var stylesheet = new Uri("https://layout.example/app.css");
        var request = Application("<link rel='stylesheet' href='app.css'><button id='hidden'>Hidden</button><button id='below'>Below</button>");
        request.Resources = new[] {
            HtmlRuntimeResource.FromText(stylesheet,
                "body{margin:0}button{display:block;width:100px;height:30px}#hidden{display:none}#below{margin-top:700px}",
                "text/css; charset=utf-8")
        };
        await using var session = await Runtime().OpenTrustedAsync(request);

        Assert.False((await session.Locator("#hidden").InspectAsync()).IsVisible);
        HtmlRuntimeElementState below = await session.Locator("#below").InspectAsync();
        Assert.True(below.IsVisible);
        Assert.False(below.IsInViewport);
        Assert.True(below.BoundingBox!.Y > 600D);
    }

    [Fact]
    public async Task ExternalStylesheetImportsParticipateWithinTheConfiguredDepth() {
        var root = new Uri("https://layout.example/app.css");
        var nested = new Uri("https://layout.example/nested.css");
        var deepest = new Uri("https://layout.example/deepest.css");
        var request = Application("<link rel='stylesheet' href='app.css'><button id='target'>Target</button>");
        request.MaxStylesheetImportDepth = 1;
        request.Resources = new[] {
            HtmlRuntimeResource.FromText(root, "@import url('nested.css');body{margin:0}", "text/css; charset=utf-8"),
            HtmlRuntimeResource.FromText(nested, "@import url('deepest.css');#target{width:90px;height:35px}", "text/css; charset=utf-8"),
            HtmlRuntimeResource.FromText(deepest, "#target{display:none}", "text/css; charset=utf-8")
        };
        await using var shallow = await Runtime().OpenTrustedAsync(request);
        Assert.True((await shallow.Locator("#target").InspectAsync()).IsVisible);

        request.MaxStylesheetImportDepth = 2;
        await using var deep = await Runtime().OpenTrustedAsync(request);
        Assert.False((await deep.Locator("#target").InspectAsync()).IsVisible);
    }

    [Fact]
    public async Task ImportedStylesheetCssomMutationsDriveInteractionLayout() {
        var root = new Uri("https://layout.example/app.css");
        var nested = new Uri("https://layout.example/nested.css");
        var request = Application("<link rel='stylesheet' href='app.css'><button id='target'>Target</button>");
        request.Resources = new[] {
            HtmlRuntimeResource.FromText(root, "@import url('nested.css');body{margin:0}", "text/css; charset=utf-8"),
            HtmlRuntimeResource.FromText(nested, "#target{width:90px;height:35px}", "text/css; charset=utf-8")
        };
        await using var session = await Runtime().OpenTrustedAsync(request);
        var target = session.Locator("#target");
        Assert.True((await target.InspectAsync()).IsVisible);

        await session.ExecuteAsync("const imported=document.styleSheets[0].cssRules[0].styleSheet;imported.insertRule('#target{display:none}',imported.cssRules.length)");

        Assert.False((await target.InspectAsync()).IsVisible);

        await session.ExecuteAsync("document.styleSheets[0].cssRules[0].styleSheet.disabled=true");
        Assert.True((await target.InspectAsync()).IsVisible);
    }

    [Fact]
    public async Task DuplicateImportInstancesRetainIndependentDisabledState() {
        var shared = new Uri("https://layout.example/shared.css");
        var request = Application("<style>@import url('shared.css');@import url('shared.css');</style><button id='target'>Target</button>");
        request.Resources = new[] {
            HtmlRuntimeResource.FromText(shared, "#target{display:none}", "text/css; charset=utf-8")
        };
        await using var session = await Runtime().OpenTrustedAsync(request);
        var target = session.Locator("#target");
        Assert.False((await target.InspectAsync()).IsVisible);

        await session.ExecuteAsync("document.styleSheets[0].cssRules[1].styleSheet.disabled=true");
        Assert.False((await target.InspectAsync()).IsVisible);

        await session.ExecuteAsync("document.styleSheets[0].cssRules[0].styleSheet.disabled=true");
        Assert.True((await target.InspectAsync()).IsVisible);
    }

    [Fact]
    public async Task LiveCssomRulesAndDisabledStateDriveInteractionLayout() {
        await using var session = await Runtime().OpenTrustedAsync(Application(
            "<style>body{margin:0}</style><button id='target'>Target</button>"));
        var target = session.Locator("#target");
        Assert.True((await target.InspectAsync()).IsVisible);

        await session.ExecuteAsync("document.styleSheets[0].insertRule('#target{display:none}',1)");
        Assert.False((await target.InspectAsync()).IsVisible);

        await session.ExecuteAsync("document.styleSheets[0].deleteRule(1)");
        Assert.True((await target.InspectAsync()).IsVisible);

        await session.ExecuteAsync("document.styleSheets[0].insertRule('#target{display:none}',1);document.styleSheets[0].disabled=true");
        Assert.True((await target.InspectAsync()).IsVisible);
    }

    [Fact]
    public async Task DisabledAndAlternateExternalStylesheetsStayInactiveForActionability() {
        var disabled = new Uri("https://layout.example/disabled.css");
        var alternate = new Uri("https://layout.example/alternate.css");
        var request = Application("<link rel='stylesheet' href='disabled.css'><link rel='alternate stylesheet' href='alternate.css'><button id='target'>Visible</button>");
        request.Resources = new[] {
            HtmlRuntimeResource.FromText(disabled, "#target{display:none}", "text/css; charset=utf-8"),
            HtmlRuntimeResource.FromText(alternate, "#target{display:none}", "text/css; charset=utf-8")
        };
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.ExecuteAsync("document.querySelector('link').disabled=true");

        var disabledState = await session.EvaluateAsync("(()=>{const link=document.querySelector('link');return {link:link.disabled,attribute:link.hasAttribute('disabled'),sheet:link.sheet&&link.sheet.disabled}})()");
        Assert.True(disabledState.GetProperty("link").GetBoolean(), disabledState.ToString());

        HtmlRuntimeElementState target = await session.Locator("#target").InspectAsync();

        Assert.True(target.IsVisible);
        Assert.NotNull(target.BoundingBox);
    }

    [Fact]
    public async Task TopmostPaintedElementGatesPointerActionsAtTheTargetCenter() {
        const string html = """
            <style>
              body{margin:0}
              #target,#cover{position:absolute;left:20px;top:20px;width:120px;height:40px}
              #target{z-index:1}
              #cover{z-index:2;background:#fff}
            </style>
            <button id='target' onclick='document.body.dataset.clicked="yes"'>Target</button>
            <div id='cover'>Cover</div>
            """;
        await using var session = await Runtime().OpenTrustedAsync(Application(html));
        var target = session.Locator("#target");

        HtmlRuntimeElementState covered = await target.InspectAsync();
        Assert.True(covered.AcceptsPointerEvents);
        Assert.False(covered.ReceivesPointerAtCenter);
        HtmlAutomationResult rejected = await session.AutomateAsync(new() {
            Query = HtmlLocatorQuery.Css("#target"),
            Action = HtmlAutomationAction.Click,
            WaitForReady = false
        });
        Assert.Equal(HtmlAutomationStatus.NotReady, rejected.Status);
        Assert.Equal(string.Empty, (await session.EvaluateAsync("document.body.dataset.clicked??''")).GetString());

        await session.ExecuteAsync("document.querySelector('#cover').style.pointerEvents='none'");
        Assert.True((await target.InspectAsync()).ReceivesPointerAtCenter);
        await target.ClickAsync();
        Assert.Equal("yes", (await session.EvaluateAsync("document.body.dataset.clicked")).GetString());
    }

    [Fact]
    public async Task FixedOverlayGatesPointerActionsAfterAutomaticScrolling() {
        const string html = """
            <style>
              body{margin:0}
              #cover{position:fixed;inset:0;z-index:2;background:#fff}
              #target{display:block;margin-top:700px;width:120px;height:40px}
            </style>
            <div id='cover'>Cover</div>
            <button id='target' onpointermove='document.body.dataset.pointer="yes"' onclick='document.body.dataset.clicked="yes"'>Target</button>
            """;
        await using var session = await Runtime().OpenTrustedAsync(Application(html));
        var target = session.Locator("#target");
        Assert.False((await target.InspectAsync()).IsInViewport);

        HtmlAutomationResult rejected = await session.AutomateAsync(new() {
            Query = HtmlLocatorQuery.Css("#target"),
            Action = HtmlAutomationAction.Click,
            WaitForReady = false
        });

        Assert.Equal(HtmlAutomationStatus.NotReady, rejected.Status);
        Assert.Equal("", (await session.EvaluateAsync("document.body.dataset.pointer??''")).GetString());
        Assert.Equal("", (await session.EvaluateAsync("document.body.dataset.clicked??''")).GetString());
    }

    [Fact]
    public async Task ScrolledStickyLayoutIsConservativelyIneligibleForPointerActions() {
        const string html = """
            <style>
              body{margin:0}
              #sticky{position:sticky;top:0;height:30px;background:#fff}
              #target{display:block;margin-top:700px;width:120px;height:40px}
            </style>
            <div id='sticky'>Sticky</div>
            <button id='target' onclick='document.body.dataset.clicked="yes"'>Target</button>
            """;
        await using var session = await Runtime().OpenTrustedAsync(Application(html));
        await session.ExecuteAsync("scrollTo(0,700)");

        HtmlRuntimeElementState target = await session.Locator("#target").InspectAsync();
        HtmlAutomationResult rejected = await session.AutomateAsync(new() {
            Query = HtmlLocatorQuery.Css("#target"),
            Action = HtmlAutomationAction.Click,
            WaitForReady = false
        });

        Assert.False(target.ReceivesPointerAtCenter);
        Assert.Equal(HtmlAutomationStatus.NotReady, rejected.Status);
        Assert.Equal("", (await session.EvaluateAsync("document.body.dataset.clicked??''")).GetString());
    }
}
