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
            new HtmlAutomationRequest { Query = HtmlLocatorQuery.Css("button"), Action = HtmlAutomationAction.Press, Value = "Enter", WaitForReady = false }
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

        HtmlRuntimeElementState target = await session.Locator("#target").InspectAsync();

        Assert.True(target.IsVisible);
        Assert.NotNull(target.BoundingBox);
    }
}
