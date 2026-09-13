using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeApplicationTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task WindowGlobalsAndEventReceiversHaveOneIdentity() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<script>var library={loaded:true};window.receivers=[];window.addEventListener('probe',function(e){receivers.push(this===window,e.target===window,e.currentTarget===window);});</script>"
        });
        await session.ExecuteAsync("dispatchEvent(new Event('probe'));document.body.innerHTML='<p>Fragment</p>'");
        var result = await session.EvaluateAsync("[window===globalThis,self===window,parent===window,top===window,frames===window,document.defaultView===window,window.library===library,...receivers]");
        Assert.Equal(10, result.GetArrayLength());
        Assert.All(result.EnumerateArray(), value => Assert.True(value.GetBoolean()));
    }

    [Fact]
    public async Task EventPathsRetainElementIdentityAndUseTheGlobalWindow() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Html = "<button id='target'>Click</button>" });
        await session.ExecuteAsync("window.paths=[];addEventListener('identity',e=>paths.push(Array.isArray(e.composedPath())&&e.composedPath()[0]===window));dispatchEvent(new Event('identity'));const target=document.querySelector('#target');target.addEventListener('click',e=>paths.push(e.composedPath()[0]===target&&e.composedPath().includes(window)));target.click()");
        await session.WaitForAsync("paths.length===2");
        Assert.All((await session.EvaluateAsync("paths")).EnumerateArray(), value => Assert.True(value.GetBoolean()));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EventPathSnapshotSurvivesListenerTreeMutationAndClearsAfterDispatch(bool nativeClick) {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Html = "<section id='parent'><button id='target'>Click</button></section>" });
        await session.ExecuteAsync("window.path=[];const target=document.querySelector('#target');const parent=target.parentNode;target.addEventListener('click',e=>{window.lastEvent=e;target.remove();path=e.composedPath();const copy=e.composedPath();copy.length=0});" + (nativeClick ? "target.click()" : "target.dispatchEvent(new Event('click',{bubbles:true}))"));
        await session.WaitForAsync("path.length>0");
        var result = await session.EvaluateAsync("({path:path.map(value=>value===window?'WINDOW':value.nodeName),target:path[0]===target,parent:path[1]===parent,after:lastEvent.composedPath().length})");
        Assert.Equal(new[] { "BUTTON", "SECTION", "BODY", "HTML", "#document", "WINDOW" }, result.GetProperty("path").EnumerateArray().Select(value => value.GetString()));
        Assert.True(result.GetProperty("target").GetBoolean());
        Assert.True(result.GetProperty("parent").GetBoolean());
        Assert.Equal(0, result.GetProperty("after").GetInt32());
    }

    [Fact]
    public async Task TimerCallbacksUseTheGlobalWindowAndRetainArguments() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        await session.ExecuteAsync("window.results={};setTimeout(function(a,b){results.timeout=[this===window,a,b]},1,'first',42);const handle=setInterval(function(value){clearInterval(handle);results.interval=[this===window,value]},5,'interval')");
        await session.WaitForAsync("!!results.timeout && !!results.interval");
        var values = await session.EvaluateAsync("results");
        Assert.True(values.GetProperty("timeout")[0].GetBoolean());
        Assert.Equal("first", values.GetProperty("timeout")[1].GetString());
        Assert.Equal(42, values.GetProperty("timeout")[2].GetInt32());
        Assert.True(values.GetProperty("interval")[0].GetBoolean());
        Assert.Equal("interval", values.GetProperty("interval")[1].GetString());
    }

    [Fact]
    public async Task ClearingAnIntervalSuppressesAlreadyQueuedCallbacks() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        await session.ExecuteAsync("window.ticks=0;window.checked=false;const timer=setInterval(()=>{ticks++;clearTimeout(timer);setTimeout(()=>checked=true,30)},1);const until=Date.now()+60;while(Date.now()<until){}");
        await session.WaitForAsync("checked===true");
        Assert.Equal(1, (await session.EvaluateAsync("ticks")).GetInt32());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-10)]
    public async Task ZeroAndNegativeIntervalDelaysYieldToTheEventLoop(int delay) {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        await session.ExecuteAsync("window.ran=false;const timer=setInterval(()=>{clearInterval(timer);ran=true}," + delay + ")");
        await session.WaitForAsync("ran===true");
    }

    [Theory]
    [InlineData("observer.observe(document.body,{})")]
    [InlineData("observer.observe(document.body,{attributes:false,attributeFilter:[]})")]
    [InlineData("observer.observe(document.body,{characterData:false,characterDataOldValue:true})")]
    [InlineData("observer.observe({}, {childList:true})")]
    public async Task InvalidObservationIsCatchableWithoutTerminatingTheSession(string observe) {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        await session.ExecuteAsync("const observer=new MutationObserver(()=>{});window.error='';try{" + observe + "}catch(e){error=e.name}");
        Assert.Equal("TypeError", (await session.EvaluateAsync("error")).GetString());
    }

    [Fact]
    public async Task ObserverCallbackFailuresFailTheSession() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(async () => {
            await session.ExecuteAsync("new MutationObserver(()=>{throw new Error('observer failure')}).observe(document.body,{attributes:true});document.body.setAttribute('data-change','yes')");
            await session.WaitForAsync("false");
        });
        Assert.Contains("observer failure", failure.Message);
    }

    [Fact]
    public async Task ObserversDeliverArraysTargetsOldValuesAndStableCallbackIdentity() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Html = "<div id='target' data-value='old'><span>Text</span></div>" });
        await session.ExecuteAsync("""
            window.records=[];
            const target=document.querySelector('#target');
            const observer=new MutationObserver(function(batch, instance){
                if(this!==observer||instance!==observer||!Array.isArray(batch))throw new Error('observer identity');
                records.push(...batch.map(record=>({type:record.type,target:record.target===target,old:record.oldValue,name:record.attributeName,added:record.addedNodes.length})));
            });
            observer.observe(target,{attributes:true,attributeOldValue:true,childList:true});
            target.setAttribute('data-value','new');
            target.appendChild(document.createElement('p'));
            """);
        await session.WaitForAsync("records.length===2");
        var result = await session.EvaluateAsync("records");
        Assert.Equal("attributes", result[0].GetProperty("type").GetString());
        Assert.Equal("old", result[0].GetProperty("old").GetString());
        Assert.Equal("data-value", result[0].GetProperty("name").GetString());
        Assert.True(result[0].GetProperty("target").GetBoolean());
        Assert.Equal("childList", result[1].GetProperty("type").GetString());
        Assert.Equal(1, result[1].GetProperty("added").GetInt32());
    }

    [Fact]
    public async Task TakingRecordsAndDisconnectingSuppressDeliveryAndAllowReobservation() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Html = "<p id='target'>Text</p>" });
        await session.ExecuteAsync("""
            window.calls=0;
            const target=document.querySelector('#target');
            const observer=new MutationObserver(()=>calls++);
            observer.observe(target,{attributes:true,attributeFilter:['data-value']});
            target.setAttribute('class','ignored');target.setAttribute('data-value','one');
            window.taken=observer.takeRecords().map(record=>record.attributeName);
            target.setAttribute('data-value','two');observer.disconnect();
            """);
        Assert.Equal("data-value", (await session.EvaluateAsync("taken"))[0].GetString());
        Assert.Equal(0, (await session.EvaluateAsync("calls")).GetInt32());
        await session.ExecuteAsync("observer.observe(target,{attributes:true});target.setAttribute('data-value','three')");
        await session.WaitForAsync("calls===1");
    }

    [Fact]
    public async Task ReusingAnObserverAfterDisconnectDoesNotRestoreOldTargets() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Html = "<p id='a'></p><p id='b'></p>" });
        await session.ExecuteAsync("window.targets=[];const observer=new MutationObserver(function(records,instance){if(this!==observer||instance!==observer)throw new Error('observer identity');targets.push(...records.map(record=>record.target.id))});const a=document.querySelector('#a'),b=document.querySelector('#b');observer.observe(a,{attributes:true});a.setAttribute('data-before','discarded');observer.disconnect();observer.disconnect();observer.observe(b,{attributes:true});a.setAttribute('data-after','ignored');b.setAttribute('data-after','observed')");
        await session.WaitForAsync("targets.length>0");
        var targets = await session.EvaluateAsync("targets");
        Assert.Equal(1, targets.GetArrayLength());
        Assert.Equal("b", targets[0].GetString());
    }

    [Theory]
    [InlineData("{get attributes(){observer.disconnect();return true}}")]
    [InlineData("{attributes:true,attributeFilter:{*[Symbol.iterator](){observer.disconnect();yield 'data-change'}}}")]
    public async Task DisconnectDuringOptionConversionCannotLeaveAnOrphanedObservation(string options) {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        await session.ExecuteAsync("window.calls=0;window.drained=false;const observer=new MutationObserver(()=>calls++);observer.observe(document.body," + options + ");observer.disconnect();document.body.setAttribute('data-change','yes');setTimeout(()=>drained=true,20)");
        await session.WaitForAsync("drained===true");
        Assert.Equal(0, (await session.EvaluateAsync("calls")).GetInt32());
    }

    [Fact]
    public async Task MicrotasksSharePromiseOrderAndReportCallbackErrors() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        await session.ExecuteAsync("window.order=[];queueMicrotask(()=>{order.push(1);queueMicrotask(()=>order.push(4))});Promise.resolve().then(()=>order.push(2));queueMicrotask(()=>order.push(3))");
        Assert.Equal("1,2,3,4", (await session.EvaluateAsync("order.join(',')")).GetString());
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.ExecuteAsync("queueMicrotask(()=>{throw new Error('microtask failure')})"));
        Assert.Contains("microtask failure", error.Message);
    }

    [Fact]
    public async Task StoragePersistsAcrossCommandsAndSupportsMethodsAndNamedProperties() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        await session.ExecuteAsync("localStorage.setItem('count',1);localStorage.name='Report';localStorage.setItem('getItem','collision');sessionStorage.name='Other'");
        var result = await session.EvaluateAsync("({count:localStorage.getItem('count'),name:localStorage.name,keys:Object.keys(localStorage),length:localStorage.length,collision:localStorage.getItem('getItem'),method:typeof localStorage.getItem,other:sessionStorage.name,brand:localStorage instanceof Storage})");
        Assert.Equal("1", result.GetProperty("count").GetString());
        Assert.Equal("Report", result.GetProperty("name").GetString());
        Assert.Equal("Other", result.GetProperty("other").GetString());
        Assert.Equal("collision", result.GetProperty("collision").GetString());
        Assert.Equal("function", result.GetProperty("method").GetString());
        Assert.Equal(3, result.GetProperty("length").GetInt32());
        Assert.Equal(2, result.GetProperty("keys").GetArrayLength());
        Assert.True(result.GetProperty("brand").GetBoolean());
        await session.ExecuteAsync("delete localStorage.name;localStorage.removeItem('count');localStorage.clear()");
        Assert.Equal(0, (await session.EvaluateAsync("localStorage.length")).GetInt32());
        Assert.Null((await session.EvaluateAsync("localStorage.key(0)")).GetString());
        await using var other = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        Assert.Equal(0, (await other.EvaluateAsync("sessionStorage.length+localStorage.length")).GetInt32());
    }

    [Fact]
    public async Task StorageQuotaChecksBothNamesAndValuesAndPreservesStateAfterFailure() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { MaxStorageCharacters = 8 });
        await session.ExecuteAsync("localStorage.setItem('key','value');window.failure='';try{localStorage.key='longer'}catch(e){failure=e.name}");
        Assert.Equal("QuotaExceededError", (await session.EvaluateAsync("failure")).GetString());
        Assert.Equal("value", (await session.EvaluateAsync("localStorage.getItem('key')")).GetString());
        await session.ExecuteAsync("localStorage.setItem('key','');localStorage.other='';localStorage.removeItem('key');localStorage.other='yes'");
        Assert.Equal("yes", (await session.EvaluateAsync("localStorage.other")).GetString());
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => Runtime().OpenTrustedAsync(new HtmlScriptRequest { MaxStorageCharacters = 0 }));
    }
}
