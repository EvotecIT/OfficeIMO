using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeScriptLifecycleTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
    private static readonly Uri Page = new("https://lifecycle.example/index.html");

    [Fact]
    public async Task CurrentScriptTracksPreparedClassicExecutionAndIsNullOtherwise() {
        const string dynamicScript = """
            const dynamic=document.createElement('script');
            dynamic.id='dynamic';
            dynamic.textContent="window.dynamicCurrent=document.currentScript&&document.currentScript.id";
            document.head.append(dynamic);
            """;
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = """
                <script id='inline'>window.inlineCurrent=document.currentScript&&document.currentScript.id</script>
                <script id='external' src='/classic.js'></script>
                <script id='module' type='module'>window.moduleCurrent=document.currentScript===null</script>
                """,
            Resources = new[] {
                HtmlRuntimeResource.FromText(new Uri(Page, "/classic.js"),
                    "window.externalCurrent=document.currentScript&&document.currentScript.id", "text/javascript")
            }
        });

        await session.ExecuteAsync(dynamicScript);
        await session.WaitForAsync("window.dynamicCurrent==='dynamic'");

        Assert.True((await session.EvaluateAsync("inlineCurrent==='inline' && externalCurrent==='external' && moduleCurrent===true && dynamicCurrent==='dynamic' && document.currentScript===null")).GetBoolean());
    }

    [Fact]
    public async Task DomInsertedInlineScriptRunsOnTheCurrentScriptStack() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = """
                <script id='outer'>
                window.order=['outer:'+document.currentScript.id];
                const inner=document.createElement('script');
                inner.id='inner';
                inner.textContent="order.push('inner:'+document.currentScript.id)";
                document.head.append(inner);
                order.push('restored:'+document.currentScript.id);
                </script>
                """
        });

        Assert.Equal("outer:outer,inner:inner,restored:outer", (await session.EvaluateAsync("order.join(',')")).GetString());
        Assert.True((await session.EvaluateAsync("document.currentScript===null")).GetBoolean());
    }

    [Fact]
    public async Task DynamicClassicScriptsWithAsyncDisabledExecuteInInsertionOrder() {
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/first.js") {
                await Task.Delay(250, token);
                return RuntimeHttpFixture.Reply.Text("order.push('first')", "text/javascript");
            }
            return RuntimeHttpFixture.Reply.Text("order.push('second')", "text/javascript");
        });
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = """
                <script>
                window.order=[];
                const first=document.createElement('script');
                const second=document.createElement('script');
                window.defaultDynamicAsync=first.async;
                first.async=false;
                second.async=false;
                window.disabledDynamicAsync=!first.async&&!second.async;
                first.src='/first.js';
                second.src='/second.js';
                document.head.append(first,second);
                </script>
                """
        });

        Assert.True((await session.EvaluateAsync("defaultDynamicAsync===true && disabledDynamicAsync===true")).GetBoolean());
        Assert.Equal("first,second", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task DynamicClassicScriptsDefaultToAsyncCompletionOrder() {
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/first.js") {
                await Task.Delay(250, token);
                return RuntimeHttpFixture.Reply.Text("order.push('first')", "text/javascript");
            }
            return RuntimeHttpFixture.Reply.Text("order.push('second')", "text/javascript");
        });
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = """
                <script>
                window.order=[];
                const first=document.createElement('script');
                const second=document.createElement('script');
                window.dynamicAsync=first.async&&second.async;
                first.src='/first.js';
                second.src='/second.js';
                document.head.append(first,second);
                </script>
                """
        });

        Assert.True((await session.EvaluateAsync("dynamicAsync")).GetBoolean());
        Assert.Equal("second,first", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task DocumentWriteReentersTheParserAndRestoresCurrentScript() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = """
                <p id='before'>Before</p>
                <script id='outer'>
                window.order=['outer:'+document.currentScript.id];
                document.write("<strong id='written'>Written</strong><script id='inner'>order.push('inner:'+document.currentScript.id)<\/script>");
                order.push('visible:'+!!document.querySelector('#written'));
                order.push('restored:'+document.currentScript.id);
                </script>
                <p id='after'>After</p>
                """
        });

        Assert.True((await session.EvaluateAsync("!!document.querySelector('#written') && !!document.querySelector('#after') && document.currentScript===null")).GetBoolean());
        Assert.Equal("outer:outer,inner:inner,visible:true,restored:outer", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task DocumentWriteYieldsAtExternalScriptAndParserResumesInOrder() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Resources = new[] {
                HtmlRuntimeResource.FromText(new Uri(Page, "/inner.js"),
                    "order.push('inner:'+document.currentScript.id+':'+!!document.querySelector('#written-after'))", "text/javascript")
            },
            Html = """
                <script id='outer'>
                window.order=['outer'];
                document.write("<script id='inner' src='/inner.js'><\/script><b id='written-after'>Written</b>");
                order.push('outer-after:'+!!document.querySelector('#written-after'));
                </script>
                <script>order.push('parser-after:'+!!document.querySelector('#written-after'))</script>
                """
        });

        Assert.Equal("outer,outer-after:false,inner:inner:false,parser-after:true", (await session.EvaluateAsync("order.join(',')")).GetString());
        Assert.True((await session.EvaluateAsync("document.currentScript===null")).GetBoolean());
    }

    [Fact]
    public async Task SequentialWrittenExternalScriptsRemainParserBlockingAndOrdered() {
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/first.js") await Task.Delay(250, token);
            return RuntimeHttpFixture.Reply.Text(path == "/first.js" ? "order.push('first')" : "order.push('second')", "text/javascript");
        });
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = """
                <script>
                window.order=[];
                document.write("<script src='/first.js'><\/script>");
                document.write("<script src='/second.js'><\/script><b id='written-tail'>Tail</b>");
                order.push('outer');
                </script>
                <script>order.push('after:'+!!document.querySelector('#written-tail'))</script>
                """
        });

        Assert.Equal("outer,first,second,after:true", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task NestedDocumentWriteConsumesNestedAndOuterTailsBeforeReturning() {
        string padding = new('x', 2048);
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Html = $$"""
                <script>
                window.order=[];
                document.write("<script>document.write('<b id=\"nested-tail\">{{padding}}</b>')<\/script><i id=\"outer-tail\">Outer</i>");
                order.push('returned:'+!!document.querySelector('#nested-tail')+':'+!!document.querySelector('#outer-tail'));
                </script>
                """
        });

        Assert.Equal("returned:true:true", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task ParserCreatedStylesheetBlocksFollowingClassicScript() {
        var styleStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseStyle = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var scriptObserved = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/style.css") {
                styleStarted.TrySetResult();
                await releaseStyle.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("body{color:green}", "text/css");
            }
            if (path == "/script-observed") scriptObserved.TrySetResult();
            return RuntimeHttpFixture.Reply.Text("ok");
        });

        var opening = Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = "<link rel='stylesheet' href='/style.css'><script>fetch('/script-observed')</script>"
        });
        await styleStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Task winner = await Task.WhenAny(scriptObserved.Task, Task.Delay(250));
        releaseStyle.TrySetResult();
        await using var session = await opening;
        await scriptObserved.Task.WaitAsync(TimeSpan.FromSeconds(5));

        Assert.NotSame(scriptObserved.Task, winner);
    }

    [Theory]
    [InlineData("rel='alternate stylesheet' title='alternate'")]
    [InlineData("media='print'")]
    [InlineData("media='(min-width: 2000px)'")]
    public async Task InactiveParserStylesheetDoesNotBlockFollowingClassicScript(string attributes) {
        var styleStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseStyle = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var scriptObserved = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/style.css") {
                styleStarted.TrySetResult();
                await releaseStyle.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("body{color:green}", "text/css");
            }
            if (path == "/script-observed") {
                scriptObserved.TrySetResult();
                releaseStyle.TrySetResult();
            }
            return RuntimeHttpFixture.Reply.Text("ok");
        });

        string relation = attributes.Contains("rel=", StringComparison.Ordinal) ? string.Empty : "rel='stylesheet' ";
        var opening = Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = "<link " + relation + attributes + " href='/style.css'><script>fetch('/script-observed')</script>"
        });
        await styleStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        try {
            await scriptObserved.Task.WaitAsync(TimeSpan.FromSeconds(2));
        } finally {
            releaseStyle.TrySetResult();
        }
        await using var session = await opening;
    }

    [Fact]
    public async Task InitiallyDisabledParserStylesheetDoesNotFetchOrBlock() {
        int styleRequests = 0;
        var scriptObserved = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture((path, _) => {
            if (path == "/style.css") Interlocked.Increment(ref styleRequests);
            if (path == "/script-observed") scriptObserved.TrySetResult();
            return Task.FromResult(RuntimeHttpFixture.Reply.Text("ok", path.EndsWith(".css", StringComparison.Ordinal) ? "text/css" : "text/plain"));
        });

        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = "<link rel='stylesheet' disabled href='/style.css'><script>fetch('/script-observed')</script>"
        });
        await scriptObserved.Task.WaitAsync(TimeSpan.FromSeconds(5));

        Assert.Equal(0, Volatile.Read(ref styleRequests));
    }

    [Fact]
    public async Task DisablingParserStylesheetReleasesAWaitingClassicScript() {
        var styleStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseStyle = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var scriptObserved = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/style.css") {
                styleStarted.TrySetResult();
                await releaseStyle.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("body{color:green}", "text/css");
            }
            if (path == "/script-observed") {
                scriptObserved.TrySetResult();
                releaseStyle.TrySetResult();
            }
            return RuntimeHttpFixture.Reply.Text("ok");
        });

        var opening = Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Resources = new[] {
                HtmlRuntimeResource.FromText(new Uri(server.Origin, "/disable.js"),
                    "document.querySelector('#sheet').disabled=true;window.disabledRan=true", "text/javascript")
            },
            Html = "<link id='sheet' rel='stylesheet' href='/style.css'><script async src='/disable.js'></script><script>fetch('/script-observed')</script>"
        });
        await styleStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        try {
            await scriptObserved.Task.WaitAsync(TimeSpan.FromSeconds(2));
        } finally {
            releaseStyle.TrySetResult();
        }
        await using var session = await opening;

        Assert.True((await session.EvaluateAsync("disabledRan===true && document.querySelector('#sheet').disabled===true")).GetBoolean());
    }

    [Fact]
    public async Task DomInsertedStylesheetDoesNotBlockFollowingParserScript() {
        var styleStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseStyle = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var scriptObserved = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/style.css") {
                styleStarted.TrySetResult();
                await releaseStyle.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("body{color:green}", "text/css");
            }
            if (path == "/script-observed") {
                scriptObserved.TrySetResult();
                releaseStyle.TrySetResult();
            }
            return RuntimeHttpFixture.Reply.Text("ok");
        });

        var opening = Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = "<script>const link=document.createElement('link');link.rel='stylesheet';link.href='/style.css';document.head.append(link)</script><script>fetch('/script-observed')</script>"
        });
        await styleStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        try {
            await scriptObserved.Task.WaitAsync(TimeSpan.FromSeconds(2));
        } finally {
            releaseStyle.TrySetResult();
        }
        await using var session = await opening;
    }

    [Fact]
    public async Task ParserCreatedInlineStyleImportBlocksFollowingClassicScript() {
        var styleStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseStyle = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var scriptObserved = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/import.css") {
                styleStarted.TrySetResult();
                await releaseStyle.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("body{color:green}", "text/css");
            }
            if (path == "/script-observed") scriptObserved.TrySetResult();
            return RuntimeHttpFixture.Reply.Text("ok");
        });

        var opening = Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = "<style>@import url('/import.css');</style><script>fetch('/script-observed')</script>"
        });
        await styleStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Task winner = await Task.WhenAny(scriptObserved.Task, Task.Delay(250));
        releaseStyle.TrySetResult();
        await using var session = await opening;
        await scriptObserved.Task.WaitAsync(TimeSpan.FromSeconds(5));

        Assert.NotSame(scriptObserved.Task, winner);
    }

    [Fact]
    public async Task ReplacingParserStylesheetHrefByCaseRetiresTheOldBlockerAndWaitsForTheReplacement() {
        var oldStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var newStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseOld = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseNew = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var scriptObserved = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/Theme.css") {
                oldStarted.TrySetResult();
                await releaseOld.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("body{color:rgb(9,8,7)}", "text/css");
            }
            if (path == "/theme.css") {
                newStarted.TrySetResult();
                await releaseNew.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("body{color:rgb(1,2,3)}", "text/css");
            }
            if (path == "/observed") scriptObserved.TrySetResult();
            return RuntimeHttpFixture.Reply.Text("ok");
        });

        var opening = Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Resources = new[] { HtmlRuntimeResource.FromText(new Uri(server.Origin, "/replace.js"),
                "document.querySelector('#sheet').href='/theme.css'", "text/javascript") },
            Html = "<link id='sheet' rel='stylesheet' href='/Theme.css'><script async src='/replace.js'></script><script>fetch('/observed')</script>"
        });
        await oldStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        await newStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Task winner = await Task.WhenAny(scriptObserved.Task, Task.Delay(250));
        releaseNew.TrySetResult();
        try {
            await using var session = await opening.WaitAsync(TimeSpan.FromSeconds(5));
            await scriptObserved.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.NotSame(scriptObserved.Task, winner);
            Assert.Contains("1", (await session.EvaluateAsync("getComputedStyle(document.body).color")).GetString());
        } finally {
            releaseOld.TrySetResult();
        }
    }

    [Fact]
    public async Task ReplacingParserStyleContentRetiresTheOldImportAndCannotRestoreItsSheet() {
        var oldStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var newStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseOld = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseNew = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var scriptObserved = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/old.css") {
                oldStarted.TrySetResult();
                await releaseOld.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("body{color:rgb(9,8,7)}", "text/css");
            }
            if (path == "/new.css") {
                newStarted.TrySetResult();
                await releaseNew.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("body{color:rgb(1,2,3)}", "text/css");
            }
            if (path == "/observed") scriptObserved.TrySetResult();
            return RuntimeHttpFixture.Reply.Text("ok");
        });

        var opening = Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Resources = new[] { HtmlRuntimeResource.FromText(new Uri(server.Origin, "/replace.js"),
                "const sheet=document.querySelector('#sheet');sheet.textContent=\"@import url('/new.css');\";window.immediateReplacementRuleCount=sheet.sheet?sheet.sheet.cssRules.length:-1", "text/javascript") },
            Html = "<style id='sheet'>@import url('/old.css');</style><script async src='/replace.js'></script><script>fetch('/observed')</script>"
        });
        await oldStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        await newStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Task winner = await Task.WhenAny(scriptObserved.Task, Task.Delay(250));
        releaseNew.TrySetResult();
        try {
            await using var session = await opening.WaitAsync(TimeSpan.FromSeconds(5));
            await scriptObserved.Task.WaitAsync(TimeSpan.FromSeconds(5));
            releaseOld.TrySetResult();
            await Task.Delay(100);
            Assert.NotSame(scriptObserved.Task, winner);
            Assert.Equal(-1, (await session.EvaluateAsync("immediateReplacementRuleCount")).GetInt32());
            var sheet = await session.EvaluateAsync("(()=>{const rule=document.querySelector('#sheet').sheet.cssRules[0];return {href:rule.href,rules:rule.styleSheet.cssRules.length}})()");
            Assert.EndsWith("/new.css", sheet.GetProperty("href").GetString(), StringComparison.Ordinal);
            Assert.Equal(1, sheet.GetProperty("rules").GetInt32());
        } finally {
            releaseOld.TrySetResult();
        }
    }

    [Fact]
    public async Task CancellationWhileAStylesheetBlocksTheParserDoesNotStrandOpening() {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (_, token) => {
            started.TrySetResult();
            await Task.Delay(Timeout.Infinite, token);
            return RuntimeHttpFixture.Reply.Text("", "text/css");
        });
        using var cancellation = new CancellationTokenSource();
        Task opening = Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = "<link rel='stylesheet' href='/slow.css'><script>window.executed=true</script>"
        }, cancellation.Token);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();

        var error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => opening.WaitAsync(TimeSpan.FromSeconds(5)));
        Assert.Equal(cancellation.Token, error.CancellationToken);
    }

    [Fact]
    public async Task BeforeScriptExecuteCancellationSkipsExternalExecutionAndContinuesParsing() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Resources = new[] { HtmlRuntimeResource.FromText(new Uri(Page, "/cancelled.js"),
                "window.cancelledScriptExecuted=true", "text/javascript") },
            Html = """
                <script>document.addEventListener('beforescriptexecute',event=>{if(event.target.id==='cancelled')event.preventDefault()},true)</script>
                <script id='cancelled' src='/cancelled.js'></script>
                <script>window.parserContinued=true</script>
                """
        });

        Assert.True((await session.EvaluateAsync("parserContinued===true && typeof cancelledScriptExecuted==='undefined'")).GetBoolean());
    }

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
