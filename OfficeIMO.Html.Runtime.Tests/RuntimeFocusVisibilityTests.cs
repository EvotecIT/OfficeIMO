using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeFocusVisibilityTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task FocusEligibilityDoesNotRequirePaintedBoundsOrPointerReception() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = "<input id='start'><span id='empty' tabindex='0'></span><input id='transparent' style='opacity:0;pointer-events:none'><input id='end'>"
        });

        await session.Locator("#start").PressAsync("Tab");
        Assert.Equal("empty", (await session.EvaluateAsync("document.activeElement.id")).GetString());
        await session.Locator("#empty").PressAsync("Tab");
        Assert.Equal("transparent", (await session.EvaluateAsync("document.activeElement.id")).GetString());
        await session.ExecuteAsync("document.querySelector('#empty').focus()");
        Assert.Equal("empty", (await session.EvaluateAsync("document.activeElement.id")).GetString());
    }

    [Fact]
    public async Task FocusDependentStylesRetainLiveFocusDuringCloneMeasurement() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = "<style>#group:not(:focus-within) #second{display:none}#first:focus{pointer-events:none}</style><div id='group'><input id='first'><input id='second'></div><input id='end'>"
        });

        await session.ExecuteAsync("document.querySelector('#first').focus()");
        Assert.Equal("first", (await session.EvaluateAsync("document.activeElement.id")).GetString());
        await session.Locator("#first").PressAsync("Tab");
        Assert.Equal("second", (await session.EvaluateAsync("document.activeElement.id")).GetString());
        await session.Locator("#second").PressAsync("Tab");
        Assert.Equal("end", (await session.EvaluateAsync("document.activeElement.id")).GetString());
        await session.ExecuteAsync("document.querySelector('#second').focus()");
        Assert.Equal("end", (await session.EvaluateAsync("document.activeElement.id")).GetString());
    }

    [Fact]
    public async Task TabSkipsCssHiddenControlsAndKeepsOffscreenFocusableControls() {
        var runtime = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
        await using var session = await runtime.OpenTrustedAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = """
                <style>#display{display:none}#visibility{visibility:hidden}#offscreen{position:absolute;top:1200px}</style>
                <input id='start'><input id='display'><input id='visibility'><input id='offscreen'><input id='end'>
                """
        });
        await session.Locator("#start").PressAsync("Tab");
        Assert.Equal("offscreen", (await session.EvaluateAsync("document.activeElement.id")).GetString());
        await session.Locator("#offscreen").PressAsync("Tab");
        Assert.Equal("end", (await session.EvaluateAsync("document.activeElement.id")).GetString());
        await session.ExecuteAsync("document.querySelector('#display').focus();document.querySelector('#visibility').focus();");
        Assert.Equal("end", (await session.EvaluateAsync("document.activeElement.id")).GetString());
        await session.ExecuteAsync("document.querySelector('#end').style.display='none';");
        Assert.Equal("BODY", (await session.EvaluateAsync("document.activeElement.tagName")).GetString());
    }
}
