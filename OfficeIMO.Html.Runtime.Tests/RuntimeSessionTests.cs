using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeSessionTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task RetainsLiveGlobalsEventsAndIndependentCapturesAcrossCommands() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<button id='add'>Add</button><p id='count'>0</p><script>let count=0;document.querySelector('#add').addEventListener('click',()=>{count++;document.querySelector('#count').textContent=String(count);});</script>"
        });
        await session.ExecuteAsync("document.querySelector('#add').click()");
        var first = await session.CaptureAsync();
        await Task.WhenAll(Enumerable.Range(0, 20).Select(_ => session.ExecuteAsync("document.querySelector('#add').click()")));
        var state = await session.EvaluateAsync("({count:count,text:document.querySelector('#count').textContent,values:[true,null,'a\\nb']})");
        Assert.Equal(21, state.GetProperty("count").GetInt32());
        Assert.Equal("21", state.GetProperty("text").GetString());
        Assert.Equal("a\nb", state.GetProperty("values")[2].GetString());
        var second = await session.CaptureAsync();
        await session.DisposeAsync();
        await session.DisposeAsync();
        Assert.Equal("1", first.Document.QuerySelector("#count")!.TextContent);
        Assert.Equal("21", second.Document.QuerySelector("#count")!.TextContent);
        Assert.NotEqual(first.Document.SnapshotId, second.Document.SnapshotId);
        await Assert.ThrowsAsync<ObjectDisposedException>(() => session.ExecuteAsync("count++"));
    }

    [Fact]
    public async Task CancelledQueuedCommandDoesNotAffectTheActiveWaitOrSession() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Html = "<p>Ready</p>" });
        await session.ExecuteAsync("window.value=0;setTimeout(()=>{window.ready=true;},300)");
        Task active = session.WaitForAsync("window.ready===true");
        using var queuedCancellation = new CancellationTokenSource();
        Task queued = session.ExecuteAsync("window.value=99", queuedCancellation.Token);
        queuedCancellation.Cancel();
        var failure = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => queued);
        Assert.Equal(queuedCancellation.Token, failure.CancellationToken);
        await active;
        Assert.Equal(0, (await session.EvaluateAsync("window.value")).GetInt32());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ActiveTimeoutOrCancellationTerminatesTheSession(bool cancel) {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Timeout = TimeSpan.FromSeconds(10) });
        using var cancellation = new CancellationTokenSource();
        Task active = session.ExecuteAsync("while(true){}", cancellation.Token);
        if (cancel) {
            cancellation.Cancel();
            var failure = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => active);
            Assert.Equal(cancellation.Token, failure.CancellationToken);
        } else await Assert.ThrowsAsync<TimeoutException>(() => active);
        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.EvaluateAsync("1"));
    }

    [Fact]
    public async Task DisposalInterruptsAnActiveCommandAndCompletesBeforeItsDeadline() {
        var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Timeout = TimeSpan.FromSeconds(30) });
        Task active = session.ExecuteAsync("while(true){}");
        await session.DisposeAsync().AsTask().WaitAsync(TimeSpan.FromSeconds(5));
        await Assert.ThrowsAsync<ObjectDisposedException>(() => active);
    }

    [Fact]
    public async Task LifetimeExpiryIncludesIdleTime() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { SessionTimeout = TimeSpan.FromSeconds(2) });
        await Task.Delay(TimeSpan.FromSeconds(2));
        await Assert.ThrowsAsync<TimeoutException>(() => session.EvaluateAsync("1"));
    }

    [Fact]
    public async Task InputValidationDoesNotTerminateAnUnusedCommand() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { MaxInputCharacters = 128 });
        await Assert.ThrowsAsync<ArgumentException>(() => session.ExecuteAsync(new string(' ', 129)));
        Assert.Equal(2, (await session.EvaluateAsync("1+1")).GetInt32());
    }

    [Theory]
    [InlineData("undefined")]
    [InlineData("(()=>{const a={};a.self=a;return a;})()")]
    public async Task UnserializableEvaluationFailsAndTerminatesTheSession(string expression) {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.EvaluateAsync(expression));
        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.CaptureAsync());
    }

    [Theory]
    [InlineData("throw new Error('session failure')")]
    [InlineData("document.body.addEventListener('click',()=>{throw new Error('session failure')});document.body.click()")]
    [InlineData("setTimeout(()=>{window.ready=true;throw new Error('session failure')},10)")]
    [InlineData("Promise.resolve().then(()=>{throw new Error('session failure')})")]
    [InlineData("document.body.onclick=()=>{window.ready=true;throw new Error('session failure')};document.body.click()")]
    [InlineData("document.body.setAttribute('onclick',\"window.ready=true;throw new Error('session failure')\");document.body.click()")]
    public async Task ScriptFailuresDoNotPublishSuccessfulSnapshots(string script) {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(async () => {
            await session.ExecuteAsync(script);
            await session.CaptureAsync("window.ready===true");
        });
        Assert.Contains("session failure", failure.Message);
    }

    [Fact]
    public async Task TracksInlineRejectionsButAllowsHandledRejectionsAndCaughtCallbackErrors() {
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<script>Promise.reject(new Error('inline rejection'))</script>"
        }));
        Assert.Contains("inline rejection", failure.Message);
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<p id='value'>pending</p><script>Promise.reject(new Error('handled')).catch(()=>document.querySelector('#value').textContent='handled');document.body.addEventListener('click',()=>{try{throw new Error('caught')}catch(e){window.caught=true;}});</script>"
        });
        await session.ExecuteAsync("document.body.click()");
        await session.WaitForAsync("window.caught===true");
        Assert.Equal("handled", (await session.CaptureAsync()).Document.QuerySelector("#value")!.TextContent);
    }

    [Fact]
    public async Task ListenerWrappersPreserveIdentityAcrossPrototypesAndThisBinding() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<body><script>window.calls=0;window.listener=function(){if(this!==document.body)throw new Error('wrong this');window.calls++;};EventTarget.prototype.addEventListener.call(document.body,'click',window.listener);</script>"
        });
        await session.ExecuteAsync("document.body.click()");
        await session.WaitForAsync("window.calls===1");
        await session.ExecuteAsync("document.body.removeEventListener('click',window.listener);document.body.click()");
        Assert.Equal(1, (await session.EvaluateAsync("window.calls")).GetInt32());
    }

    [Fact]
    public async Task EventHandlersSupportReplacementRemovalIdentityAndCancellation() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<button id='button' onclick='window.calls++;return false'>Click</button><div id='container'></div><script>window.calls=0;</script>"
        });
        Assert.False((await session.EvaluateAsync("document.querySelector('#button').dispatchEvent(new Event('click',{cancelable:true}))")).GetBoolean());
        Assert.Equal(1, (await session.EvaluateAsync("window.calls")).GetInt32());
        await session.ExecuteAsync("const b=document.querySelector('#button');b.setAttribute('onclick','window.calls+=10');b.click()");
        await session.WaitForAsync("window.calls===11");
        await session.ExecuteAsync("b.removeAttribute('onclick');b.click()");
        Assert.Equal(11, (await session.EvaluateAsync("window.calls")).GetInt32());
        await session.ExecuteAsync("window.handler=function(){if(this!==b)throw new Error('wrong receiver');window.calls+=100;};b.onclick=window.handler");
        Assert.True((await session.EvaluateAsync("b.onclick===window.handler")).GetBoolean());
        await session.ExecuteAsync("b.click()");
        await session.WaitForAsync("window.calls===111");
        await session.ExecuteAsync("b.onclick=null;b.click();document.querySelector('#container').innerHTML='<button id=late onclick=\"window.calls++\">Late</button>';document.querySelector('#late').click()");
        await session.WaitForAsync("window.calls===112");
        Assert.Null((await session.EvaluateAsync("b.onclick")).GetString());
    }

    [Fact]
    public async Task ListenersSupportObjectsOnceDuplicateSuppressionAndRemovalAfterFragmentParsing() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Html = "<button id='b'>Click</button><div id='fragment'></div>" });
        await session.ExecuteAsync("window.calls=0;const b=document.querySelector('#b');const listener={handleEvent(e){if(this!==listener||e.currentTarget!==b)throw new Error('wrong receiver');window.calls++;}};b.addEventListener('click',listener);b.addEventListener('click',listener);document.querySelector('#fragment').innerHTML='<p>Parsed</p>';b.click()");
        await session.WaitForAsync("window.calls===1");
        await session.ExecuteAsync("b.removeEventListener('click',listener);b.click()");
        Assert.Equal(1, (await session.EvaluateAsync("window.calls")).GetInt32());
        await session.ExecuteAsync("b.addEventListener('click',listener,{once:true});b.dispatchEvent(new Event('click'));b.dispatchEvent(new Event('click'))");
        Assert.Equal(2, (await session.EvaluateAsync("window.calls")).GetInt32());
    }

    [Fact]
    public async Task PendingPromiseRejectionsAreBoundedEvenWhenHandledLaterInTheSameTurn() {
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            MaxPendingPromiseRejections = 2,
            Html = "<script>const pending=[Promise.reject(1),Promise.reject(2),Promise.reject(3)];pending.forEach(p=>p.catch(()=>{}));</script>"
        }));
        Assert.Contains("tracking budget", failure.Message);
    }

    [Theory]
    [InlineData("window.")]
    [InlineData("")]
    public async Task WindowListenersUseTheSameRegistrationAndErrorContract(string receiver) {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        await session.ExecuteAsync($"window.calls=0;window.listener=()=>window.calls++;{receiver}addEventListener('test',window.listener);{receiver}dispatchEvent(new Event('test'));");
        Assert.Equal(1, (await session.EvaluateAsync("window.calls")).GetInt32());
        await session.ExecuteAsync($"{receiver}removeEventListener('test',window.listener);{receiver}dispatchEvent(new Event('test'));");
        Assert.Equal(1, (await session.EvaluateAsync("window.calls")).GetInt32());
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.ExecuteAsync($"{receiver}addEventListener('test',()=>{{throw new Error('window failure')}});{receiver}dispatchEvent(new Event('test'))"));
        Assert.Contains("window failure", failure.Message);
    }
}
