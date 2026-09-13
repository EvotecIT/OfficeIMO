using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeInputAutomationTests {
    private static readonly Uri Start = new("https://input.example/start");
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
    private static HtmlScriptRequest Application(string html) => new() {
        Profile = HtmlRuntimeProfile.WebApplicationV1,
        Html = html,
        DocumentUrl = Start,
        ViewportWidth = 320,
        ViewportHeight = 160
    };

    [Fact]
    public async Task HoverAndClickDispatchSelectedPrimaryPointerSequenceAtLayoutCoordinates() {
        const string html = """
            <style>body{margin:0}button{display:block;margin-top:500px;width:120px;height:40px}</style>
            <button id='target'>Target</button>
            <script>
              window.events=[];
              const target=document.querySelector('#target');
              for(const name of ['pointerover','pointerenter','mouseover','mouseenter','pointermove','mousemove','pointerdown','mousedown','pointerup','mouseup','click'])
                target.addEventListener(name,e=>events.push({name,x:e.clientX,y:e.clientY,mouse:e instanceof MouseEvent}));
              target.addEventListener('mousedown',e=>e.preventDefault());
            </script>
            """;
        await using var session = await Runtime().OpenTrustedAsync(Application(html));
        var target = session.Locator("#target");

        await target.HoverAsync();
        await target.ClickAsync();

        var result = await session.EvaluateAsync("({names:events.map(e=>e.name).join(','),coordinates:events.every(e=>e.x>=0&&e.x<=innerWidth&&e.y>=0&&e.y<=innerHeight),mouse:events.every(e=>e.mouse),focused:document.activeElement===target})");
        Assert.Equal("pointerover,pointerenter,mouseover,mouseenter,pointermove,mousemove,pointermove,mousemove,pointerdown,mousedown,pointerup,mouseup,click", result.GetProperty("names").GetString());
        Assert.True(result.GetProperty("coordinates").GetBoolean());
        Assert.True(result.GetProperty("mouse").GetBoolean());
        Assert.False(result.GetProperty("focused").GetBoolean());
    }

    [Fact]
    public async Task PressDispatchesKeyboardEventsAndSelectedEndOfValueEditing() {
        const string html = """
            <style>input,button{display:block;width:120px;height:32px}</style>
            <input id='query' value='a'><button id='next'>Next</button>
            <script>
              window.events=[];
              window.buttonEvents=[];
              const query=document.querySelector('#query');
              const next=document.querySelector('#next');
              for(const name of ['keydown','beforeinput','input','keyup','change','blur'])
                query.addEventListener(name,e=>events.push(name+':'+(e.key??e.data??'')));
              query.addEventListener('keydown',e=>{if(e.key==='x')e.preventDefault()});
              for(const name of ['keydown','keyup','click'])next.addEventListener(name,()=>buttonEvents.push(name));
            </script>
            """;
        await using var session = await Runtime().OpenTrustedAsync(Application(html));
        var query = session.Locator("#query");

        await query.PressAsync("b");
        await query.PressAsync("x");
        await query.PressAsync("Backspace");
        await query.PressAsync("c");
        await query.PressAsync("Tab");
        await session.ExecuteAsync("buttonEvents=[]");
        await session.Locator("#next").PressAsync("Space");

        var result = await session.EvaluateAsync("({value:query.value,events:events.join(','),next:document.activeElement.id,buttonEvents:buttonEvents.join(',')})");
        Assert.Equal("ac", result.GetProperty("value").GetString());
        Assert.Contains("keydown:b,beforeinput:b,input:b,keyup:b", result.GetProperty("events").GetString());
        Assert.Contains("keydown:x,keyup:x", result.GetProperty("events").GetString());
        Assert.DoesNotContain("input:x", result.GetProperty("events").GetString());
        Assert.Contains("change:,blur:", result.GetProperty("events").GetString());
        Assert.Equal("next", result.GetProperty("next").GetString());
        Assert.Equal("keydown,keyup,click", result.GetProperty("buttonEvents").GetString());
    }

    [Fact]
    public async Task PointerHandlersCannotActivateAStaleTarget() {
        const string html = """
            <style>button{display:block;width:120px;height:32px}</style>
            <button id='target'>Target</button>
            <script>
              window.events=[];
              const target=document.querySelector('#target');
              for(const name of ['pointerover','pointerdown','mousedown','pointerup','mouseup','click'])
                target.addEventListener(name,()=>events.push(name));
              target.addEventListener('pointerover',()=>target.style.pointerEvents='none');
            </script>
            """;
        await using var session = await Runtime().OpenTrustedAsync(Application(html));

        HtmlAutomationResult result = await session.AutomateAsync(new() {
            Query = HtmlLocatorQuery.Css("#target"),
            Action = HtmlAutomationAction.Click,
            WaitForReady = false
        });

        Assert.Equal(HtmlAutomationStatus.Rejected, result.Status);
        Assert.Equal("pointerover", (await session.EvaluateAsync("events.join(',')")).GetString());
    }

    [Fact]
    public async Task PointerDepartureHandlersCannotReachAnInvalidatedArrivalTarget() {
        const string html = """
            <style>button{display:block;width:120px;height:32px}</style>
            <button id='first'>First</button><button id='target'>Target</button>
            <script>
              window.arrivals=[];
              const first=document.querySelector('#first'),target=document.querySelector('#target');
              first.addEventListener('pointerout',()=>target.style.pointerEvents='none');
              target.addEventListener('pointerover',()=>arrivals.push('pointerover'));
            </script>
            """;
        await using var session = await Runtime().OpenTrustedAsync(Application(html));
        await session.Locator("#first").HoverAsync();

        HtmlAutomationResult result = await session.AutomateAsync(new() {
            Query = HtmlLocatorQuery.Css("#target"),
            Action = HtmlAutomationAction.Click,
            WaitForReady = false
        });

        Assert.Equal(HtmlAutomationStatus.Rejected, result.Status);
        Assert.Equal(string.Empty, (await session.EvaluateAsync("arrivals.join(',')")).GetString());
    }

    [Fact]
    public async Task SpaceKeyupCannotActivateAControlThatChangesType() {
        const string html = """
            <style>input{display:block;width:120px;height:32px}</style>
            <input id='target' type='checkbox'>
            <script>
              window.events=[];
              const target=document.querySelector('#target');
              for(const name of ['keydown','keyup','click'])target.addEventListener(name,()=>events.push(name));
              target.addEventListener('keyup',()=>target.type='text');
            </script>
            """;
        await using var session = await Runtime().OpenTrustedAsync(Application(html));

        HtmlAutomationResult result = await session.AutomateAsync(new() {
            Query = HtmlLocatorQuery.Css("#target"),
            Action = HtmlAutomationAction.Press,
            Value = "Space",
            WaitForReady = false
        });

        Assert.Equal(HtmlAutomationStatus.Rejected, result.Status);
        Assert.Equal("keydown,keyup", (await session.EvaluateAsync("events.join(',')")).GetString());
        Assert.False((await session.EvaluateAsync("target.checked")).GetBoolean());
    }

    [Fact]
    public async Task EnterOnTextInputRunsTheOwnedGetFormNavigationDefault() {
        var resultUrl = new Uri("https://input.example/results?q=ab");
        const string html = """
            <style>input,button{display:block;width:120px;height:32px}</style>
            <form action='/results' method='get'><input id='query' name='q' value='a'><button>Search</button></form>
            """;
        var request = Application(html);
        request.Resources = new[] { HtmlRuntimeResource.FromText(resultUrl, "<h1 id='result'>Result</h1>", "text/html; charset=utf-8") };
        await using var session = await Runtime().OpenTrustedAsync(request);
        var query = session.Locator("#query");

        await query.PressAsync("b");
        await query.PressAsync("Enter");

        Assert.Equal(resultUrl, (await session.CaptureAsync()).DocumentUrl);
        Assert.Equal("Result", (await session.Locator("#result").InspectAsync()).Text);
    }
}
