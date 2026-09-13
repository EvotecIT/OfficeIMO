using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeScriptLifecycleTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
    private static readonly Uri Page = new("https://lifecycle.example/index.html");

    [Fact]
    public async Task ModuleScriptCanJoinAnImportWhileItsSourceWaitsForATimer() {
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        int requests = 0;
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("ok")));
        server.RespondToRequest = async (request, token) => {
            if (request.Path == "/shared.js") {
                if (Interlocked.Increment(ref requests) == 1) await release.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("window.executions=(window.executions||0)+1;export const value=42", "text/javascript");
            }
            if (request.Path == "/release") release.TrySetResult();
            return RuntimeHttpFixture.Reply.Text("ok");
        };
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true }, Timeout = TimeSpan.FromSeconds(10),
            Html = "<script>window.joined=import('/shared.js').then(m=>window.value=m.value);setTimeout(()=>fetch('/release'),1000)</script><script type='module' src='/shared.js'></script>"
        });
        Assert.True((await session.EvaluateAsync("executions===1 && value===42")).GetBoolean());
        Assert.Equal(2, requests);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task WindowCaptureCanStopPropagationToDocumentOrElement(bool documentEvent) {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = """
                <body><p id='target'>Ready</p><script>
                window.order=[];
                const type=TYPE;
                window.addEventListener(type,e=>{order.push('capture');e.stopPropagation()},true);
                window.addEventListener(type,()=>order.push('same-window'),true);
                window.addEventListener(type,()=>order.push('bubble'));
                document.addEventListener(type,()=>order.push('document'));
                document.body.addEventListener(type,()=>order.push('body'));
                document.querySelector('#target').addEventListener(type,()=>order.push('target'));
                DISPATCH
                </script>
                """.Replace("TYPE", documentEvent ? "'DOMContentLoaded'" : "'probe'")
                    .Replace("DISPATCH", documentEvent ? "" : "document.querySelector('#target').dispatchEvent(new Event('probe',{bubbles:true}))")
        });
        Assert.Equal("capture,same-window", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task ADeferredModuleAwaitDoesNotHoldTheFollowingModule() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = "<script>window.order=[];document.addEventListener('DOMContentLoaded',()=>order.push('dcl'))</script><script type='module'>order.push('one');await new Promise(resolve=>document.addEventListener('DOMContentLoaded',resolve));order.push('one:done')</script><script type='module'>order.push('two')</script>"
        });
        Assert.Equal("one,two,dcl,one:done", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task LoadIncludesScriptsInsertedByAnotherLoadBlocker() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = "<script>window.order=[];document.addEventListener('DOMContentLoaded',()=>{order.push('dcl');const script=document.createElement('script');script.src='/first.js';document.head.append(script)});window.addEventListener('load',()=>order.push('load'))</script>",
            Resources = new[] {
                HtmlRuntimeResource.FromText(new Uri(Page, "/first.js"), "order.push('first');const next=document.createElement('script');next.src='/second.js';document.head.append(next)", "text/javascript"),
                HtmlRuntimeResource.FromText(new Uri(Page, "/second.js"), "order.push('second')", "text/javascript")
            }
        });
        Assert.Equal("dcl,first,second,load", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ModulePreparationRetainsTypeSourceAndBaseWhenMarkupChanges(bool external) {
        const string module = "import {value} from './dep.js';window.result=value;window.moduleUrl=import.meta.url";
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = "<base href='/assets/'>" + (external ? "<script id='module' type='module' src='app.js'></script>" : "<script id='module' type='module'>" + module + "</script>") +
                "<script>const target=document.querySelector('#module');target.src='/wrong.js';target.type='text/plain';document.querySelector('base').href='/other/'</script>",
            Resources = new[] {
                HtmlRuntimeResource.FromText(new Uri(Page, "/assets/app.js"), module, "text/javascript"),
                HtmlRuntimeResource.FromText(new Uri(Page, "/assets/dep.js"), "export const value='prepared'", "text/javascript")
            }
        });
        Assert.Equal("prepared", (await session.EvaluateAsync("result")).GetString());
        Assert.Equal(external ? "https://lifecycle.example/assets/app.js" : "https://lifecycle.example/assets/", (await session.EvaluateAsync("moduleUrl")).GetString());
    }

    [Fact]
    public async Task ReadinessEventsUseOneDocumentBubbleAndOneWindowLoad() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = """
                <script>
                window.order=[];
                document.addEventListener('readystatechange',()=>order.push(document.readyState));
                document.addEventListener('DOMContentLoaded',e=>order.push('document:'+e.bubbles+':'+(e.target===document)));
                window.addEventListener('DOMContentLoaded',e=>order.push('window:'+(e.target===document)));
                window.addEventListener('load',()=>order.push('load'));
                </script>
                """
        });
        Assert.Equal("interactive,document:true:true,window:true,complete,load", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task AsyncExecutionDoesNotRewindOrRaceTheActiveParser() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = "<body><script>window.executions=0</script><script async src='/async.js'></script>" + string.Concat(Enumerable.Range(0, 1000).Select(i => "<p data-row='" + i + "'>Row</p>")),
            Resources = new[] { HtmlRuntimeResource.FromText(new Uri(Page, "/async.js"), "executions++;document.body.setAttribute('data-async','done')", "text/javascript") }
        });
        Assert.True((await session.EvaluateAsync("executions===1 && document.body.getAttribute('data-async')==='done' && document.querySelectorAll('[data-row]').length===1000 && document.querySelector('[data-row=\"999\"]')!==null")).GetBoolean());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ParserModulesSeeTheCompletedTreeAndRunBeforeDomContentLoaded(bool external) {
        const string module = "window.order.push('module:'+document.readyState+':'+!!document.querySelector('#later'));document.querySelector('#later').textContent='Ready'";
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = "<script>window.order=[];document.addEventListener('DOMContentLoaded',()=>order.push('dcl'))</script>" +
                (external ? "<script type='module' src='/app.js'></script>" : "<script type='module'>" + module + "</script>") +
                "<p id='later'>Loading</p><script>order.push('classic:'+document.readyState)</script>",
            Resources = new[] { HtmlRuntimeResource.FromText(new Uri(Page, "/app.js"), module, "text/javascript") }
        });
        Assert.Equal("classic:loading,module:interactive:true,dcl", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task ModuleAwaitCanDependOnDomContentLoadedAndWindowLoad() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page, Timeout = TimeSpan.FromSeconds(10),
            Html = """
                <script>window.order=[];document.addEventListener('DOMContentLoaded',()=>order.push('dcl'));window.addEventListener('load',()=>order.push('load'))</script>
                <script type='module'>
                order.push('start');
                const loaded=new Promise(resolve=>window.addEventListener('load',resolve,{once:true}));
                await new Promise(resolve=>document.addEventListener('DOMContentLoaded',resolve,{once:true}));
                order.push('after-dcl');await loaded;order.push('after-load');
                </script>
                """
        });
        Assert.Equal("start,dcl,after-dcl,load,after-load", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task SlowAsyncScriptDoesNotDelayDeferredScriptsOrDomContentLoadedButDelaysLoad() {
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("ok")));
        server.RespondToRequest = async (request, token) => {
            if (request.Path == "/slow.js") {
                await release.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("order.push('async')", "text/javascript");
            }
            if (request.Path == "/dcl") release.TrySetResult();
            return RuntimeHttpFixture.Reply.Text("ok");
        };
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true },
            Html = "<script>window.order=[];document.addEventListener('DOMContentLoaded',()=>{order.push('dcl');fetch('/dcl')});window.addEventListener('load',()=>order.push('load'))</script><script async src='/slow.js'></script><script defer src='/defer.js'></script>",
            Resources = new[] { HtmlRuntimeResource.FromText(new Uri(server.Origin, "/defer.js"), "order.push('defer')", "text/javascript") }
        });
        Assert.Equal("defer,dcl,async,load", (await session.EvaluateAsync("order.join(',')")).GetString());
    }
}
