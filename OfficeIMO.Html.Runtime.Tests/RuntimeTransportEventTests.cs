using System.Text;
using System.Text.Json;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeTransportEventTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task DomXhrAndAbortTargetsInvokeCaptureBeforeHandlersAndBubbleListeners() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Html = "<button></button>" });
        JsonElement result = await session.EvaluateAsync("""
            (()=>{const button=document.querySelector('button'),xhr=new XMLHttpRequest(),signal=new AbortController().signal;
                function collect(target,type,handler){const order=[];
                    target[handler]=()=>order.push('old-handler');
                    target.addEventListener(type,()=>order.push('bubble'));
                    target[handler]=()=>order.push('handler');
                    target.addEventListener(type,e=>order.push('capture'),true);
                    target.dispatchEvent(new Event(type,{bubbles:true}));return order;}
                const dom=collect(button,'click','onclick'),transport=collect(xhr,'load','onload'),abort=collect(signal,'abort','onabort');
                const stopped=[],stopping=new AbortController();
                stopping.signal.onabort=()=>stopped.push('handler');
                stopping.signal.addEventListener('abort',e=>{stopped.push('capture');e.stopPropagation();},true);
                stopping.signal.addEventListener('abort',()=>stopped.push('second-capture'),true);stopping.abort();
                const immediate=[],controller=new AbortController();
                controller.signal.onabort=()=>immediate.push('handler');
                controller.signal.addEventListener('abort',e=>{immediate.push('capture');e.stopImmediatePropagation();},true);
                controller.abort();return {dom,transport,abort,stopped,immediate};})()
            """);

        foreach (string target in new[] { "dom", "transport", "abort" })
            Assert.Equal(new[] { "capture", "handler", "bubble" }, result.GetProperty(target).EnumerateArray().Select(value => value.GetString()));
        Assert.Equal(new[] { "capture", "second-capture" }, result.GetProperty("stopped").EnumerateArray().Select(value => value.GetString()));
        Assert.Equal(new[] { "capture" }, result.GetProperty("immediate").EnumerateArray().Select(value => value.GetString()));
    }

    [Fact]
    public async Task XhrOpenFreezesItsRelativeUrlBeforeBaseAndHistoryChanges() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri("https://app.example/a/page"),
            Html = "<base href='/a/'>",
            Resources = new[] {
                new HtmlRuntimeResource(new Uri("https://app.example/a/report.txt"), Encoding.UTF8.GetBytes("original"), "text/plain"),
                new HtmlRuntimeResource(new Uri("https://app.example/b/report.txt"), Encoding.UTF8.GetBytes("changed"), "text/plain")
            }
        });
        await session.ExecuteAsync("""
            window.xhr=new XMLHttpRequest(); xhr.open('GET','report.txt');
            document.querySelector('base').href='/b/'; history.pushState(null,'','/b/page');
            xhr.onload=()=>{window.result={url:xhr.responseURL,text:xhr.responseText};window.done=true;};
            xhr.onerror=()=>{window.result={url:xhr.responseURL,text:xhr.responseText};window.done=true;};
            xhr.send();
            """);
        await session.WaitForAsync("window.done===true");
        JsonElement result = await session.EvaluateAsync("window.result");
        Assert.Equal("https://app.example/a/report.txt", result.GetProperty("url").GetString());
        Assert.Equal("original", result.GetProperty("text").GetString());
        Assert.Equal("SyntaxError", (await session.EvaluateAsync("(()=>{try{new XMLHttpRequest().open('GET','https://[');}catch(e){return e.name;}})()")).GetString());
    }

    [Fact]
    public async Task XhrAndAbortSignalShareNativeEventIdentityOrderingAndImmediatePropagation() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri("https://app.example/"),
            Resources = new[] { new HtmlRuntimeResource(new Uri("https://app.example/report"), Encoding.UTF8.GetBytes("ok"), "text/plain") }
        });
        await session.ExecuteAsync("""
            window.order=[];window.native={};window.xhr=new XMLHttpRequest();xhr.open('GET','/report');
            window.readystatechange=null;xhr.onreadystatechange=e=>{window.readystatechange={event:e instanceof Event,progress:e instanceof ProgressEvent};};
            xhr.onload=function(e){order.push('handler');native={event:e instanceof Event,progress:e instanceof ProgressEvent,
                target:e.target===xhr,current:e.currentTarget===xhr,receiver:this===xhr,trusted:e.isTrusted,
                phase:e.eventPhase,path:e.composedPath().length};window.seen=e;};
            xhr.addEventListener('load',e=>{order.push('listener');e.stopImmediatePropagation();},{once:true});
            xhr.addEventListener('load',()=>order.push('suppressed'));
            xhr.onloadend=()=>{window.done=true;};xhr.send();
            """);
        await session.WaitForAsync("window.done===true");
        JsonElement result = await session.EvaluateAsync("({order,native,readystatechange,after:seen.currentTarget===null,target:xhr instanceof EventTarget})");
        Assert.Equal(new[] { "handler", "listener" }, result.GetProperty("order").EnumerateArray().Select(value => value.GetString()));
        foreach (string flag in new[] { "event", "progress", "target", "current", "receiver", "trusted" })
            Assert.True(result.GetProperty("native").GetProperty(flag).GetBoolean(), flag);
        Assert.Equal(2, result.GetProperty("native").GetProperty("phase").GetInt32());
        Assert.Equal(1, result.GetProperty("native").GetProperty("path").GetInt32());
        Assert.True(result.GetProperty("after").GetBoolean());
        Assert.True(result.GetProperty("target").GetBoolean());
        Assert.True(result.GetProperty("readystatechange").GetProperty("event").GetBoolean());
        Assert.False(result.GetProperty("readystatechange").GetProperty("progress").GetBoolean());

        JsonElement abort = await session.EvaluateAsync("""
            (()=>{const controller=new AbortController(),signal=controller.signal,order=[];
                signal.onabort=e=>{order.push('old');};
                signal.addEventListener('abort',e=>{order.push('listener');e.stopImmediatePropagation();});
                signal.onabort=e=>{order.push('handler');window.abortEvent=e;};
                signal.addEventListener('abort',()=>order.push('suppressed'));
                controller.abort();
                const native=abortEvent instanceof Event && abortEvent.target===signal && abortEvent.isTrusted;
                signal.onabort=null;signal.onabort=()=>order.push('late');
                signal.dispatchEvent(new Event('abort'));
                return {order,native,target:signal instanceof EventTarget,after:abortEvent.currentTarget===null};})()
            """);
        Assert.Equal(new[] { "handler", "listener", "listener" }, abort.GetProperty("order").EnumerateArray().Select(value => value.GetString()));
        foreach (string flag in new[] { "native", "target", "after" }) Assert.True(abort.GetProperty(flag).GetBoolean(), flag);

        JsonElement cancellation = await session.EvaluateAsync("""
            (()=>{const xhr=new XMLHttpRequest();let trusted;
                xhr.onload=e=>{trusted=e.isTrusted;return false;};
                const event=new Event('load',{cancelable:true});
                return {result:xhr.dispatchEvent(event),prevented:event.defaultPrevented,trusted};})()
            """);
        Assert.False(cancellation.GetProperty("result").GetBoolean());
        Assert.True(cancellation.GetProperty("prevented").GetBoolean());
        Assert.False(cancellation.GetProperty("trusted").GetBoolean());
    }

    [Fact]
    public async Task TransportListenersRetainIdentityOncePassiveAndDispatchValidation() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        JsonElement result = await session.EvaluateAsync("""
            (()=>{const xhr=new XMLHttpRequest(),calls=[],listener={handleEvent(e){calls.push('removed');}};
                xhr.addEventListener('probe',listener);xhr.addEventListener('probe',listener);
                EventTarget.prototype.removeEventListener.call(xhr,'probe',listener);
                xhr.addEventListener('probe',e=>{calls.push('once');e.preventDefault();},{once:true,passive:true});
                xhr.addEventListener('probe',e=>{calls.push('ordinary');try{xhr.dispatchEvent(e);}catch(error){calls.push(error.name);}});
                const event=new Event('probe',{cancelable:true});EventTarget.prototype.dispatchEvent.call(xhr,event);xhr.dispatchEvent(event);
                let invalid;try{xhr.dispatchEvent({type:'probe'});}catch(error){invalid=error.name;}
                const limited=new XMLHttpRequest(),callback=()=>{};
                for(let i=0;i<128;i++)limited.addEventListener(String(i),callback);
                limited.addEventListener('0',callback);
                let quota;try{limited.addEventListener('extra',callback);}catch(error){quota=error.name;}
                limited.removeEventListener('0',callback);limited.addEventListener('extra',callback);
                const other=new XMLHttpRequest(),pathEvent=new Event('path'),paths=[];
                xhr.addEventListener('path',e=>paths.push(e.composedPath()[0]===xhr));
                other.addEventListener('path',e=>paths.push(e.composedPath()[0]===other));
                xhr.dispatchEvent(pathEvent);other.dispatchEvent(pathEvent);
                return {calls,paths,prevented:event.defaultPrevented,invalid,quota};})()
            """);
        Assert.Equal(new[] { "once", "ordinary", "InvalidStateError", "ordinary", "InvalidStateError" },
            result.GetProperty("calls").EnumerateArray().Select(value => value.GetString()));
        Assert.False(result.GetProperty("prevented").GetBoolean());
        Assert.Equal("TypeError", result.GetProperty("invalid").GetString());
        Assert.Equal("QuotaExceededError", result.GetProperty("quota").GetString());
        Assert.All(result.GetProperty("paths").EnumerateArray(), value => Assert.True(value.GetBoolean()));
    }

    [Fact]
    public async Task HeadersIteratorsFollowLiveSortedPairPositionsDuringMutation() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest());
        JsonElement result = await session.EvaluateAsync("""
            (()=>{const headers=new Headers([['b','2'],['c','3']]),iterator=headers.keys(),added=[];
                added.push(iterator.next().value);headers.append('a','1');added.push(...iterator);
                const removed=new Headers([['a','1'],['b','2'],['c','3']]),next=removed.keys(),deleted=[];
                deleted.push(next.next().value);removed.delete('a');deleted.push(...next);
                return {added,deleted};})()
            """);
        Assert.Equal(new[] { "b", "b", "c" }, result.GetProperty("added").EnumerateArray().Select(value => value.GetString()));
        Assert.Equal(new[] { "a", "c" }, result.GetProperty("deleted").EnumerateArray().Select(value => value.GetString()));
    }
}
