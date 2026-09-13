using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeEventHandlerContractsTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task StoppingAtTheTargetDoesNotLeakToAncestorsOrTheNextDispatch(bool immediate) {
        await using var session = await Runtime().OpenTrustedAsync(new() { Html = "<button>Go</button>" });
        await session.ExecuteAsync("""
            window.order=[];const target=document.querySelector('button');
            target.addEventListener('probe',e=>{order.push('stop');e.METHOD()},{once:true});
            target.addEventListener('probe',()=>order.push('target'));
            document.body.addEventListener('probe',()=>order.push('body'));
            window.addEventListener('probe',()=>order.push('window'));
            const event=new Event('probe',{bubbles:true});target.dispatchEvent(event);
            order.push('again');target.dispatchEvent(event);
            """.Replace("METHOD", immediate ? "stopImmediatePropagation" : "stopPropagation"));
        Assert.Equal(immediate ? "stop,again,target,body,window" : "stop,target,again,target,body,window", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task ActionPromiseJobsCompleteBeforeTheNextTypedWaitOrCapture() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<button>Go</button><p id='status'>Pending</p>",
            Scripts = new[] { "document.querySelector('button').onclick=()=>Promise.resolve().then(()=>{document.querySelector('#status').textContent='First';queueMicrotask(()=>document.querySelector('#status').textContent='Ready')})" }
        });
        await session.Locator("button").ClickAsync();
        await session.Locator("#status").WaitForTextAsync("Ready");
        Assert.Equal("Ready", (await session.CaptureAsync()).Document.QuerySelector("#status")!.TextContent);
    }

    [Fact]
    public async Task BodyWindowHandlersUseTheWindowSlotWhileOrdinaryHandlersStayLocal() {
        await using var session = await Runtime().OpenTrustedAsync(new() { Html = "<body><button>Go</button></body>" });
        await session.ExecuteAsync("window.calls=[];window.handler=function(e){calls.push(e.type+':'+(this===window))};document.body.onresize=handler;document.body.onclick=function(){calls.push('body:'+ (this===document.body))}");
        Assert.True((await session.EvaluateAsync("window.onresize===handler && document.body.onresize===handler")).GetBoolean());
        await session.ExecuteAsync("dispatchEvent(new Event('resize'));document.body.dispatchEvent(new Event('click'))");
        Assert.Equal("resize:true,body:true", (await session.EvaluateAsync("calls.join(',')")).GetString());
        await session.ExecuteAsync("window.onresize=null;document.body.setAttribute('onresize',\"calls.push('attribute')\");dispatchEvent(new Event('resize'));document.body.removeAttribute('onresize')");
        Assert.True((await session.EvaluateAsync("window.onresize===null && calls[calls.length-1]==='attribute'")).GetBoolean());
    }

    [Fact]
    public async Task UnhandledActionPromiseRejectionFailsTheActionAndSession() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<select><option>A</option><option>B</option></select>",
            Scripts = new[] { "document.querySelector('select').onchange=()=>Promise.resolve().then(()=>{throw new Error('selection failure')})" }
        });
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.Locator("select").SelectOptionsAsync(new[] { "B" }));
        Assert.Contains("selection failure", failure.Message);
        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.CaptureAsync());
    }

    [Fact]
    public async Task NonterminatingActionPromiseIsBoundedByTheCommandDeadline() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<button>Go</button>", Timeout = TimeSpan.FromSeconds(2),
            Scripts = new[] { "document.querySelector('button').onclick=()=>Promise.resolve().then(()=>{while(true){}})" }
        });
        await Assert.ThrowsAsync<TimeoutException>(() => session.Locator("button").ClickAsync());
        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.CaptureAsync());
    }

    [Fact]
    public async Task TimerPromiseHandledAtItsMicrotaskCheckpointDoesNotFailATypedWait() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<p id='status'>Pending</p>",
            Scripts = new[] { "setTimeout(()=>{const rejected=Promise.reject(new Error('handled'));queueMicrotask(()=>rejected.catch(()=>document.querySelector('#status').textContent='Handled'))},30)" }
        });
        await session.Locator("#status").WaitForTextAsync("Handled");
        Assert.Equal("Handled", (await session.CaptureAsync()).Document.QuerySelector("#status")!.TextContent);
    }

    [Fact]
    public async Task SelectPropertyHandlersRetainIdentityAndObserveOwnedActions() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<label for='choice'>Region</label><select id='choice'><option value='north'>North</option><option value='south'>South</option></select>",
            Scripts = new[] { "window.calls=[];window.focusHandler=function(e){calls.push(e.type+':'+(this===document.querySelector('select')))};const select=document.querySelector('select');select.onfocus=focusHandler;select.oninput=e=>calls.push(e.type+':'+e.target.value);select.onchange=e=>calls.push(e.type+':'+e.target.value)" }
        });
        Assert.True((await session.EvaluateAsync("document.querySelector('select').onfocus===focusHandler")).GetBoolean());
        await session.Locator(HtmlLocatorQuery.ByAccessibleName("Region")).SelectOptionsAsync(new[] { "south" });
        Assert.Equal("focus:true,input:south,change:south", (await session.EvaluateAsync("calls.join(',')")).GetString());
        await session.ExecuteAsync("document.querySelector('select').onchange=null");
        await session.Locator("select").SelectOptionsAsync(new[] { "north" });
        Assert.Equal("focus:true,input:south,change:south,input:north", (await session.EvaluateAsync("calls.join(',')")).GetString());
    }

    [Theory]
    [InlineData("button", "<button>Go</button>")]
    [InlineData("select", "<select><option>A</option></select>")]
    public async Task HandlerReplacementKeepsListenerPositionAndRemovalReleasesIt(string selector, string html) {
        await using var session = await Runtime().OpenTrustedAsync(new() { Html = html });
        await session.ExecuteAsync("window.calls=[];window.target=document.querySelector('" + selector + "');target.addEventListener('click',()=>calls.push('first'));target.onclick=()=>calls.push('old');target.addEventListener('click',()=>calls.push('last'));target.onclick=()=>calls.push('replacement')");
        await session.Locator(selector).ClickAsync();
        Assert.Equal("first,replacement,last", (await session.EvaluateAsync("calls.join(',')")).GetString());
        await session.ExecuteAsync("calls=[];target.onclick=null;target.onclick=()=>calls.push('new')");
        await session.Locator(selector).ClickAsync();
        Assert.Equal("first,last,new", (await session.EvaluateAsync("calls.join(',')")).GetString());
    }

    [Fact]
    public async Task SelectAttributesAndPrototypeAccessorsShareTheHandlerSlot() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<select onclick='window.calls++;return false'><option>A</option></select>",
            Scripts = new[] { "window.calls=0" }
        });
        Assert.False((await session.EvaluateAsync("document.querySelector('select').dispatchEvent(new Event('click',{cancelable:true}))")).GetBoolean());
        Assert.Equal(1, (await session.EvaluateAsync("calls")).GetInt32());
        await session.ExecuteAsync("window.select=document.querySelector('select');window.handler=()=>calls+=10;Object.getOwnPropertyDescriptor(HTMLElement.prototype,'onclick').set.call(select,handler)");
        Assert.True((await session.EvaluateAsync("select.onclick===handler && Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype,'onclick').get.call(select)===handler")).GetBoolean());
        await session.Locator("select").ClickAsync();
        Assert.Equal(11, (await session.EvaluateAsync("calls")).GetInt32());
        await session.ExecuteAsync("select.removeAttribute('onclick')");
        Assert.True((await session.EvaluateAsync("select.onclick===null")).GetBoolean());
    }

    [Fact]
    public async Task SelectOrdinaryPropertiesAndOptionLookupsPreserveNativeIdentity() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<select><option id='north' value='n'>North</option><option id='south' value='s'>South</option></select>"
        });
        await session.ExecuteAsync("window.select=document.querySelector('select');select.multiple=true;select.disabled=true;select.value='s';select.selectedIndex=0;select.dataset.report='active'");
        var properties = await session.EvaluateAsync("({multiple:select.multiple,disabled:select.disabled,value:select.value,index:select.selectedIndex,report:select.dataset.report})");
        Assert.True(properties.GetProperty("multiple").GetBoolean());
        Assert.True(properties.GetProperty("disabled").GetBoolean());
        Assert.Equal("n", properties.GetProperty("value").GetString());
        Assert.Equal(0, properties.GetProperty("index").GetInt32());
        Assert.Equal("active", properties.GetProperty("report").GetString());
        Assert.True((await session.EvaluateAsync("select[0]===select.options[0] && select.item(1)===document.querySelector('#south') && select.namedItem('north')===select[0] && select.options.namedItem('north')===select[0] && Array.from(select.options).length===2 && select.item(-1)===null && select.namedItem('')===null")).GetBoolean());
    }

    [Fact]
    public async Task SelectValueAndIndexAssignmentsUseOrdinalSingleSelectionAndCaptureTheSameState() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<select multiple><option value='a'>Lower</option><option value='A'>Upper</option><option value='a'>Duplicate</option></select>"
        });
        await session.ExecuteAsync("window.select=document.querySelector('select');window.changes=0;select.onchange=()=>changes++;select.value='A'");
        Assert.Equal(new[] { "A" }, (await session.Locator("select").InspectAsync()).SelectedValues);
        await session.ExecuteAsync("select.value='a'");
        Assert.Equal(new[] { "a" }, (await session.Locator("select").InspectAsync()).SelectedValues);
        Assert.Equal(0, (await session.EvaluateAsync("select.selectedIndex")).GetInt32());
        await session.ExecuteAsync("select.options.selectedIndex=2");
        Assert.Equal(2, (await session.EvaluateAsync("select.selectedIndex")).GetInt32());
        var captured = await session.CaptureAsync();
        Assert.True(captured.Document.QuerySelectorAll("option")[2].FormState!.IsSelected);
        Assert.False(captured.Document.QuerySelectorAll("option")[0].FormState!.IsSelected);
        await session.ExecuteAsync("select.selectedIndex=-1");
        Assert.Equal(string.Empty, (await session.EvaluateAsync("select.value")).GetString());
        Assert.Equal(string.Empty, (await session.Locator("select").InspectAsync()).Value);
        Assert.Empty((await session.Locator("select").InspectAsync()).SelectedValues);
        Assert.Equal(0, (await session.EvaluateAsync("changes")).GetInt32());
    }
}
