using System.Text;
using System.Text.Json;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Providers;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeFetchTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
    private static readonly Uri Origin = new("https://app.example/");
    private static async Task<JsonElement> RunAsync(string script, HtmlScriptRequest? request = null) {
        request ??= new HtmlScriptRequest { DocumentUrl = Origin };
        request.Html = "<p id='status'></p><script>(async()=>{" + script + "})().then(value=>{window.result=value;window.done=true;});</script>";
        await using var session = await Runtime().OpenTrustedAsync(request);
        await session.WaitForAsync("window.done===true");
        return await session.EvaluateAsync("window.result");
    }

    [Fact]
    public async Task SuppliedJsonUsesPromisesHeadersClonesAndIndependentBodyConsumption() {
        var resource = new HtmlRuntimeResource(new Uri(Origin, "api"), Encoding.UTF8.GetBytes("{\"total\":42}"), "application/json", headers: new Dictionary<string, string> { ["X-Report"] = "monthly", ["Set-Cookie"] = "private=value" });
        var result = await RunAsync("""
            const pending=fetch('/api'); const nativePromise=pending instanceof Promise;
            const r=await pending, clone=r.clone();
            const before=r.bodyUsed;
            const value=await r.json();
            const after=r.bodyUsed;
            let repeated, immutable;
            try { await r.text(); } catch(e) { repeated=e.name; }
            try { r.headers.set('x-report','changed'); } catch(e) { immutable=e.name; }
            return {nativePromise,global:window.fetch===fetch,response:r instanceof Response,status:r.status,ok:r.ok,
                header:r.headers.get('X-Report'),cookie:r.headers.get('set-cookie'),before,after,value,
                copied:await clone.text(),repeated,immutable};
            """, new HtmlScriptRequest { DocumentUrl = Origin, Resources = new[] { resource } });
        Assert.True(result.GetProperty("nativePromise").GetBoolean());
        Assert.True(result.GetProperty("global").GetBoolean());
        Assert.True(result.GetProperty("response").GetBoolean());
        Assert.Equal(200, result.GetProperty("status").GetInt32());
        Assert.True(result.GetProperty("ok").GetBoolean());
        Assert.Equal("monthly", result.GetProperty("header").GetString());
        Assert.Equal(JsonValueKind.Null, result.GetProperty("cookie").ValueKind);
        Assert.False(result.GetProperty("before").GetBoolean());
        Assert.True(result.GetProperty("after").GetBoolean());
        Assert.Equal(42, result.GetProperty("value").GetProperty("total").GetInt32());
        Assert.Equal("{\"total\":42}", result.GetProperty("copied").GetString());
        Assert.Equal("TypeError", result.GetProperty("repeated").GetString());
        Assert.Equal("TypeError", result.GetProperty("immutable").GetString());
    }

    [Fact]
    public async Task XmlHttpRequestUsesSharedTransportAndBufferedLifecycle() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text(
            "{\"total\":42}", "application/json", headers: "X-Report: monthly\r\nSet-Cookie: private=value\r\n")));
        var result = await RunAsync("""
            return await new Promise((resolve,reject)=>{
              const xhr=new XMLHttpRequest(), events=[];
              for(const name of ['readystatechange','loadstart','load','error','loadend'])
                xhr.addEventListener(name,()=>events.push(name+(name==='readystatechange'?':'+xhr.readyState:'')));
              xhr.open('GET','/api/report?view=summary#client');
              xhr.responseType='json';
              xhr.setRequestHeader('X-Trace','42');
              xhr.onerror=()=>reject(new Error('unexpected XMLHttpRequest error'));
              xhr.onloadend=()=>{
                let responseTextError; try { void xhr.responseText; } catch(e) { responseTextError=e.name; }
                resolve({global:window.XMLHttpRequest===XMLHttpRequest,constants:[xhr.UNSENT,XMLHttpRequest.DONE],
                  events,status:xhr.status,statusText:xhr.statusText,url:xhr.responseURL,total:xhr.response.total,
                  report:xhr.getResponseHeader('X-Report'),cookie:xhr.getResponseHeader('Set-Cookie'),
                  all:xhr.getAllResponseHeaders(),responseTextError});
              };
              xhr.send();
            });
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true } });

        Assert.True(result.GetProperty("global").GetBoolean());
        Assert.Equal(new[] { 0, 4 }, result.GetProperty("constants").EnumerateArray().Select(value => value.GetInt32()));
        Assert.Equal(new[] { "readystatechange:1", "loadstart", "readystatechange:2", "readystatechange:3", "readystatechange:4", "load", "loadend" },
            result.GetProperty("events").EnumerateArray().Select(value => value.GetString()));
        Assert.Equal(200, result.GetProperty("status").GetInt32());
        Assert.Equal("Test", result.GetProperty("statusText").GetString());
        Assert.Equal(new Uri(server.Origin, "api/report?view=summary").AbsoluteUri, result.GetProperty("url").GetString());
        Assert.Equal(42, result.GetProperty("total").GetInt32());
        Assert.Equal("monthly", result.GetProperty("report").GetString());
        Assert.Equal(JsonValueKind.Null, result.GetProperty("cookie").ValueKind);
        Assert.Contains("x-report: monthly", result.GetProperty("all").GetString(), StringComparison.Ordinal);
        Assert.DoesNotContain("set-cookie", result.GetProperty("all").GetString(), StringComparison.OrdinalIgnoreCase);
        Assert.Equal("InvalidStateError", result.GetProperty("responseTextError").GetString());
        RuntimeHttpFixture.ReceivedRequest request = Assert.Single(server.Received);
        Assert.Equal("GET", request.Method);
        Assert.Equal("/api/report?view=summary", request.Path);
        Assert.Equal("42", request.Headers["X-Trace"]);
    }

    [Fact]
    public async Task XmlHttpRequestAbortAndUnsupportedModesRemainBounded() {
        var result = await RunAsync("""
            let synchronous,credentials,timeout;
            try { new XMLHttpRequest().open('GET','/',false); } catch(e) { synchronous=e.name; }
            try { const xhr=new XMLHttpRequest(); xhr.withCredentials=true; } catch(e) { credentials=e.name; }
            try { const xhr=new XMLHttpRequest(); xhr.timeout=1; } catch(e) { timeout=e.name; }
            const aborted=await new Promise(resolve=>{
              const xhr=new XMLHttpRequest(), events=[];
              for(const name of ['readystatechange','loadstart','abort','loadend'])
                xhr.addEventListener(name,()=>events.push(name+(name==='readystatechange'?':'+xhr.readyState:'')));
              xhr.open('GET','/slow');
              xhr.onloadend=()=>resolve({events,state:xhr.readyState,status:xhr.status});
              xhr.send(); xhr.abort();
            });
            return {synchronous,credentials,timeout,aborted};
            """);

        Assert.Equal("NotSupportedError", result.GetProperty("synchronous").GetString());
        Assert.Equal("NotSupportedError", result.GetProperty("credentials").GetString());
        Assert.Equal("NotSupportedError", result.GetProperty("timeout").GetString());
        JsonElement aborted = result.GetProperty("aborted");
        Assert.Equal(new[] { "readystatechange:1", "loadstart", "readystatechange:4", "abort", "loadend" },
            aborted.GetProperty("events").EnumerateArray().Select(value => value.GetString()));
        Assert.Equal(4, aborted.GetProperty("state").GetInt32());
        Assert.Equal(0, aborted.GetProperty("status").GetInt32());
    }

    [Fact]
    public async Task XmlHttpRequestPostsBinaryAndReturnsAnArrayBuffer() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(
            new RuntimeHttpFixture.Reply(new byte[] { 1, 0, 255, 7 }, "application/octet-stream")));
        var result = await RunAsync("""
            return await new Promise((resolve,reject)=>{
              const xhr=new XMLHttpRequest();
              xhr.open('POST','/binary');
              xhr.responseType='arraybuffer';
              xhr.setRequestHeader('Content-Type','application/octet-stream');
              xhr.onload=()=>resolve({status:xhr.status,bytes:Array.from(new Uint8Array(xhr.response))});
              xhr.onerror=()=>reject(new Error('unexpected XMLHttpRequest error'));
              xhr.send(new Uint8Array([9,8,7]));
            });
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true } });

        Assert.Equal(200, result.GetProperty("status").GetInt32());
        Assert.Equal(new[] { 1, 0, 255, 7 }, result.GetProperty("bytes").EnumerateArray().Select(value => value.GetInt32()));
        RuntimeHttpFixture.ReceivedRequest request = Assert.Single(server.Received);
        Assert.Equal("POST", request.Method);
        Assert.Equal(new byte[] { 9, 8, 7 }, request.Body);
        Assert.Equal("application/octet-stream", request.Headers["Content-Type"]);
    }

    [Fact]
    public async Task XmlHttpRequestPolicyFailureDispatchesErrorAndLoadEnd() {
        var result = await RunAsync("""
            return await new Promise(resolve=>{
              const xhr=new XMLHttpRequest(), events=[];
              for(const name of ['readystatechange','loadstart','load','error','loadend'])
                xhr.addEventListener(name,()=>events.push(name+(name==='readystatechange'?':'+xhr.readyState:'')));
              xhr.open('GET','file:///private');
              xhr.onloadend=()=>resolve({events,status:xhr.status,url:xhr.responseURL,text:xhr.responseText});
              xhr.send();
            });
            """);

        Assert.Equal(new[] { "readystatechange:1", "loadstart", "readystatechange:4", "error", "loadend" },
            result.GetProperty("events").EnumerateArray().Select(value => value.GetString()));
        Assert.Equal(0, result.GetProperty("status").GetInt32());
        Assert.Equal("", result.GetProperty("url").GetString());
        Assert.Equal("", result.GetProperty("text").GetString());
    }

    [Fact]
    public async Task XmlHttpRequestSuppliedGetParticipatesInReadinessAndCapture() {
        Uri dataUrl = new(Origin, "legacy/report.json");
        HtmlScriptCapture capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = Origin,
            Html = """
                <p id="result">Loading</p>
                <script>
                  const xhr=new XMLHttpRequest();
                  xhr.open('GET','/legacy/report.json');
                  xhr.responseType='json';
                  xhr.onload=()=>document.querySelector('#result').textContent='Legacy total '+xhr.response.total;
                  xhr.send();
                </script>
                """,
            ReadyExpression = "document.querySelector('#result')?.textContent === 'Legacy total 42'",
            Resources = new[] { HtmlRuntimeResource.FromText(dataUrl, "{\"total\":42}", "application/json") }
        });

        Assert.Contains("Legacy total 42", capture.Document.Body!.TextContent, StringComparison.Ordinal);
        HtmlRuntimeResource retained = Assert.Single(capture.Resources);
        Assert.Equal(dataUrl, retained.Url);
        Assert.Equal("application/json", retained.ContentType);
    }

    [Fact]
    public async Task XmlHttpRequestReplayEligibilityRequiresAHeaderlessGet() {
        await using IHtmlRuntimeContext context = await Runtime().CreateContextAsync(new HtmlRuntimeContextOptions {
            Trace = new HtmlRuntimeTraceOptions { IncludeUrls = true }
        });
        await using IHtmlRuntimePage page = await context.OpenPageAsync(new HtmlScriptRequest {
            DocumentUrl = Origin,
            Html = """
                <script>
                  window.xhrDone=0;
                  for(const [path,header] of [['/headerless',false],['/headered',true]]) {
                    const xhr=new XMLHttpRequest();
                    xhr.open('GET',path);
                    if(header) xhr.setRequestHeader('X-Variant','private');
                    xhr.onloadend=()=>window.xhrDone++;
                    xhr.send();
                  }
                </script>
                """,
            ReadyExpression = "window.xhrDone===2"
        });

        await page.CaptureAsync();
        HtmlRuntimeEvent[] decisions = page.GetTrace().Events.Where(item =>
            item.Kind == HtmlRuntimeEventKind.Policy && item.Operation == "network-access" && item.Method == "GET").ToArray();
        Assert.Contains(decisions, item => item.Url == new Uri(Origin, "headerless") &&
            item.Decision == "network-disabled-replayable-get");
        Assert.Contains(decisions, item => item.Url == new Uri(Origin, "headered") &&
            item.Decision == "network-disabled");
    }

    [Fact]
    public async Task XmlHttpRequestReentrantLifecyclePreservesReplacementRequests() {
        await using var server = new RuntimeHttpFixture((path, _) => Task.FromResult(
            RuntimeHttpFixture.Reply.Text(path, "text/plain")));
        var result = await RunAsync("""
            const noSend=await new Promise(resolve=>{
              const xhr=new XMLHttpRequest();
              xhr.open('GET','/original-not-sent');
              xhr.onloadstart=()=>{
                xhr.open('GET','/replacement-not-sent');
                setTimeout(()=>resolve({state:xhr.readyState,status:xhr.status}),0);
              };
              xhr.send();
            });
            const loading=await new Promise((resolve,reject)=>{
              const xhr=new XMLHttpRequest(); let replaced=false,loads=0,loadEnds=0;
              xhr.addEventListener('load',()=>loads++);
              xhr.addEventListener('loadend',()=>{
                loadEnds++;
                if(replaced) resolve({state:xhr.readyState,text:xhr.responseText,loads,loadEnds});
              });
              xhr.onreadystatechange=()=>{
                if(xhr.readyState===xhr.LOADING && !replaced){
                  replaced=true;
                  xhr.open('GET','/second');
                  xhr.send();
                }
              };
              xhr.onerror=()=>reject(new Error('unexpected XMLHttpRequest error'));
              xhr.open('GET','/first');
              xhr.send();
            });
            const aborted=await new Promise((resolve,reject)=>{
              const xhr=new XMLHttpRequest();
              xhr.onabort=()=>{
                xhr.open('GET','/after-abort');
                xhr.onload=()=>resolve({state:xhr.readyState,text:xhr.responseText,status:xhr.status});
                xhr.onerror=()=>reject(new Error('replacement XMLHttpRequest failed'));
                xhr.send();
              };
              xhr.open('GET','/aborted');
              xhr.send();
              xhr.abort();
            });
            return {noSend,loading,aborted};
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true } });

        Assert.Equal(1, result.GetProperty("noSend").GetProperty("state").GetInt32());
        Assert.Equal(0, result.GetProperty("noSend").GetProperty("status").GetInt32());
        JsonElement loading = result.GetProperty("loading");
        Assert.Equal(4, loading.GetProperty("state").GetInt32());
        Assert.Equal("/second", loading.GetProperty("text").GetString());
        Assert.Equal(1, loading.GetProperty("loads").GetInt32());
        Assert.Equal(1, loading.GetProperty("loadEnds").GetInt32());
        JsonElement aborted = result.GetProperty("aborted");
        Assert.Equal(4, aborted.GetProperty("state").GetInt32());
        Assert.Equal("/after-abort", aborted.GetProperty("text").GetString());
        Assert.Equal(200, aborted.GetProperty("status").GetInt32());
        string[] requests = server.Requests.ToArray();
        Assert.DoesNotContain("/original-not-sent", requests);
        Assert.DoesNotContain("/replacement-not-sent", requests);
        Assert.Contains("/first", requests);
        Assert.Contains("/second", requests);
        Assert.Contains("/after-abort", requests);
    }

    [Fact]
    public async Task XmlHttpRequestResponseTypeSurvivesOpenAndTextIsVisibleWhileLoading() {
        await using var server = new RuntimeHttpFixture((path, _) => Task.FromResult(path == "/json"
            ? RuntimeHttpFixture.Reply.Text("{\"total\":42}", "application/json")
            : RuntimeHttpFixture.Reply.Text("plain", "text/plain")));
        var result = await RunAsync("""
            const json=await new Promise((resolve,reject)=>{
              const xhr=new XMLHttpRequest();
              xhr.responseType='json';
              xhr.open('GET','/json');
              xhr.onload=()=>resolve({type:xhr.responseType,total:xhr.response.total});
              xhr.onerror=()=>reject(new Error('unexpected XMLHttpRequest error'));
              xhr.send();
            });
            const text=await new Promise((resolve,reject)=>{
              const xhr=new XMLHttpRequest(); let during;
              xhr.open('GET','/text');
              xhr.onreadystatechange=()=>{
                if(xhr.readyState===xhr.HEADERS_RECEIVED) xhr.responseType='text';
                if(xhr.readyState===xhr.LOADING) during={response:xhr.response,responseText:xhr.responseText};
              };
              xhr.onload=()=>resolve({during,response:xhr.response,responseText:xhr.responseText});
              xhr.onerror=()=>reject(new Error('unexpected XMLHttpRequest error'));
              xhr.send();
            });
            return {json,text};
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true } });

        Assert.Equal("json", result.GetProperty("json").GetProperty("type").GetString());
        Assert.Equal(42, result.GetProperty("json").GetProperty("total").GetInt32());
        JsonElement text = result.GetProperty("text");
        Assert.Equal("plain", text.GetProperty("during").GetProperty("response").GetString());
        Assert.Equal("plain", text.GetProperty("during").GetProperty("responseText").GetString());
        Assert.Equal("plain", text.GetProperty("response").GetString());
        Assert.Equal("plain", text.GetProperty("responseText").GetString());
    }

    [Fact]
    public async Task HttpErrorsAreResponsesAndNetworkErrorsCanRenderAnApplicationErrorState() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("missing", "text/plain", 404)));
        var result = await RunAsync("""
            const r=await fetch('/missing');
            let caught; try { await fetch('file:///private'); } catch(e) { caught=e.name; document.querySelector('#status').textContent='Handled'; }
            return {status:r.status,ok:r.ok,text:await r.text(),caught,statusText:r.statusText,state:document.querySelector('#status').textContent};
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true } });
        Assert.Equal(404, result.GetProperty("status").GetInt32());
        Assert.False(result.GetProperty("ok").GetBoolean());
        Assert.Equal("missing", result.GetProperty("text").GetString());
        Assert.Equal("TypeError", result.GetProperty("caught").GetString());
        Assert.Equal("Handled", result.GetProperty("state").GetString());
        Assert.Equal("Test", result.GetProperty("statusText").GetString());
        Assert.Single(server.Requests);
    }

    [Fact]
    public async Task PostSendsUtf8AndBinaryBodiesAndFiltersForbiddenHeaders() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("ok")));
        var result = await RunAsync("""
            await fetch('/text',{method:'post',body:'żółw 🐢',headers:{Cookie:'private',Host:'wrong.example','X-Trace':'42','X-HTTP-Method-Override':'TRACE'}});
            const bytes=new Uint8Array([1,0,255,7]);
            await fetch('/bytes',{method:'PUT',body:bytes.subarray(1,3)});
            return true;
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true } });
        Assert.True(result.GetBoolean());
        var requests = server.Received.ToArray();
        Assert.Equal("POST", requests[0].Method);
        Assert.Equal(Encoding.UTF8.GetBytes("żółw 🐢"), requests[0].Body);
        Assert.Equal("text/plain;charset=UTF-8", requests[0].Headers["Content-Type"]);
        Assert.Equal("42", requests[0].Headers["X-Trace"]);
        Assert.False(requests[0].Headers.ContainsKey("Cookie"));
        Assert.False(requests[0].Headers.ContainsKey("X-HTTP-Method-Override"));
        Assert.Equal(server.Origin.Authority, requests[0].Headers["Host"]);
        Assert.Equal(new byte[] { 0, 255 }, requests[1].Body);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task CorsRequiresResponsePermissionInAdditionToTheHostAllowlist(bool grant) {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("{\"total\":42}", "application/json", headers:
            (grant ? "Access-Control-Allow-Origin: https://app.example\r\nAccess-Control-Expose-Headers: X-Public\r\n" : "") + "X-Public: yes\r\nX-Private: secret\r\nSet-Cookie: secret=value\r\n")));
        var result = await RunAsync("""
            try {
                const r=await fetch(TARGET);
                return {type:r.type,public:r.headers.get('x-public'),private:r.headers.get('x-private'),cookie:r.headers.get('set-cookie'),value:await r.json()};
            } catch(e) { return {error:e.name}; }
            """.Replace("TARGET", JsonSerializer.Serialize(server.Origin)), new HtmlScriptRequest { DocumentUrl = Origin, ResourcePolicy = new() { AllowNetwork = true, AllowedOrigins = new[] { server.Origin } } });
        Assert.Single(server.Requests);
        if (!grant) { Assert.Equal("TypeError", result.GetProperty("error").GetString()); return; }
        Assert.Equal("cors", result.GetProperty("type").GetString());
        Assert.Equal("yes", result.GetProperty("public").GetString());
        Assert.Equal(JsonValueKind.Null, result.GetProperty("private").ValueKind);
        Assert.Equal(JsonValueKind.Null, result.GetProperty("cookie").ValueKind);
        Assert.Equal(42, result.GetProperty("value").GetProperty("total").GetInt32());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task CorsPreflightRunsBeforeUnsafeRequestsAndChecksAuthorizationExplicitly(bool grant) {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("ok")));
        server.RespondToRequest = (request, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("ok", headers:
            "Access-Control-Allow-Origin: *\r\nAccess-Control-Allow-Methods: PUT\r\nAccess-Control-Allow-Headers: " + (grant ? "Authorization, X-Trace" : "*") + "\r\n"));
        var result = await RunAsync("""
            try { return {text:await (await fetch(TARGET,{method:'PUT',headers:{Authorization:'Bearer test','X-Trace':'42'},body:'data'})).text()}; }
            catch(e) { return {error:e.name}; }
            """.Replace("TARGET", JsonSerializer.Serialize(server.Origin)), new HtmlScriptRequest { DocumentUrl = Origin, ResourcePolicy = new() { AllowNetwork = true, AllowedOrigins = new[] { server.Origin } } });
        var requests = server.Received.ToArray();
        Assert.Equal("OPTIONS", requests[0].Method);
        Assert.Equal("https://app.example", requests[0].Headers["Origin"]);
        Assert.Equal("authorization,x-trace", requests[0].Headers["Access-Control-Request-Headers"]);
        if (grant) {
            Assert.Equal(2, requests.Length);
            Assert.Equal("PUT", requests[1].Method);
            Assert.Equal("data", Encoding.UTF8.GetString(requests[1].Body));
            Assert.Equal("ok", result.GetProperty("text").GetString());
        } else { Assert.Single(requests); Assert.Equal("TypeError", result.GetProperty("error").GetString()); }
    }

    [Fact]
    public async Task AbortRejectsWithTheExactReasonAndLeavesTheSessionUsable() {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (_, token) => { started.TrySetResult(); await Task.Delay(5000, token); return RuntimeHttpFixture.Reply.Text("late"); });
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { DocumentUrl = server.Origin, Html = "<p>Ready</p>", ResourcePolicy = new() { AllowNetwork = true } });
        await session.ExecuteAsync("window.controller=new AbortController();window.reason={message:'stop'};fetch('/slow',{signal:controller.signal}).catch(e=>{window.same=e===reason;window.done=true;});");
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        await session.ExecuteAsync("controller.abort(reason)");
        await session.WaitForAsync("window.done===true");
        Assert.True((await session.EvaluateAsync("window.same")).GetBoolean());
        Assert.Equal("Ready", (await session.CaptureAsync()).Document.QuerySelector("p")!.TextContent);
    }

    [Fact]
    public async Task AlreadyAbortedAndUnsupportedRequestsDoNotContactTheNetwork() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("unexpected")));
        var result = await RunAsync("""
            const failures=[];
            for(const init of [{signal:AbortSignal.abort('stop')},{method:'GET',body:'bad'},{credentials:'include'},{mode:'no-cors'},{redirect:'manual'}]){
                try { await fetch('/api',init); } catch(e) { failures.push(e==='stop'?'stop':e.name); }
            }
            return failures;
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true } });
        Assert.Equal(new[] { "stop", "TypeError", "TypeError", "TypeError", "TypeError" }, result.EnumerateArray().Select(item => item.GetString()));
        Assert.Empty(server.Requests);
    }

    [Theory]
    [InlineData(302, "GET", "")]
    [InlineData(303, "GET", "")]
    [InlineData(307, "POST", "payload")]
    public async Task RedirectsApplyMethodAndBodyRules(int status, string method, string body) {
        await using var server = new RuntimeHttpFixture((path, _) => Task.FromResult(path == "/start" ? RuntimeHttpFixture.Reply.Text("", status: status, headers: "Location: /final\r\n") : RuntimeHttpFixture.Reply.Text("done")));
        var result = await RunAsync("const r=await fetch('/start',{method:'POST',body:'payload'});return {redirected:r.redirected,url:r.url,text:await r.text()};",
            new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true } });
        var final = server.Received.Last();
        Assert.Equal(method, final.Method);
        Assert.Equal(body, Encoding.UTF8.GetString(final.Body));
        if (method == "GET") Assert.False(final.Headers.ContainsKey("Content-Type"));
        Assert.True(result.GetProperty("redirected").GetBoolean());
        Assert.Equal(new Uri(server.Origin, "final").AbsoluteUri, result.GetProperty("url").GetString());
        Assert.Equal("done", result.GetProperty("text").GetString());
    }

    [Fact]
    public async Task HeadHasNoBodyAndBinaryResponsesPreserveBytes() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(new RuntimeHttpFixture.Reply(new byte[] { 0, 1, 255 }, "application/octet-stream")));
        var result = await RunAsync("""
            const head=await fetch('/api',{method:'HEAD'}), text=await head.text(), again=await head.text();
            const binary=await (await fetch('/api')).arrayBuffer();
            return {text,again,used:head.bodyUsed,body:head.body,bytes:Array.from(new Uint8Array(binary))};
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true } });
        Assert.Equal("", result.GetProperty("text").GetString());
        Assert.Equal("", result.GetProperty("again").GetString());
        Assert.False(result.GetProperty("used").GetBoolean());
        Assert.Equal(JsonValueKind.Null, result.GetProperty("body").ValueKind);
        Assert.Equal(new[] { 0, 1, 255 }, result.GetProperty("bytes").EnumerateArray().Select(item => item.GetInt32()));
    }

    [Fact]
    public async Task RequestBodyLimitsCountUtf8BytesAndRedirectReplays() {
        await using var server = new RuntimeHttpFixture((path, _) => Task.FromResult(path == "/start" ? RuntimeHttpFixture.Reply.Text("", status: 307, headers: "Location: /final\r\n") : RuntimeHttpFixture.Reply.Text("ok")));
        var result = await RunAsync("""
            let oversized, replay;
            try { await fetch('/large',{method:'POST',body:'🐢🐢'}); } catch(e) { oversized=String(e); }
            try { await fetch('/start',{method:'POST',body:'12345'}); } catch(e) { replay=String(e); }
            return {oversized,replay};
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true, MaxRequestBytes = 5, MaxTotalRequestBytes = 9 } });
        Assert.Contains("byte budget", result.GetProperty("oversized").GetString());
        Assert.Contains("byte budget", result.GetProperty("replay").GetString());
        Assert.Equal("/start", Assert.Single(server.Requests));
    }

    [Fact]
    public async Task LiveBaseUriResolvesFetchAndSuccessfulGetsRemainInIndependentCaptures() {
        var source = HtmlRuntimeResource.FromText(new Uri(Origin, "data/value.json"), "{\"count\":7}", "application/json");
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = Origin, Html = "<head><base href='/data/'></head><p>Pending</p>", Resources = new[] { source }
        });
        var before = await session.CaptureAsync();
        await session.ExecuteAsync("fetch('value.json#ignored').then(r=>r.json()).then(value=>{document.querySelector('p').textContent=String(value.count);window.done=true});");
        var after = await session.CaptureAsync("window.done===true");
        await session.DisposeAsync();
        Assert.Empty(before.Resources);
        Assert.Equal("7", after.Document.QuerySelector("p")!.TextContent);
        var resource = Assert.Single(after.Resources);
        Assert.Equal(source.Url, resource.FinalUrl);
        Assert.Equal("application/json", resource.Headers["Content-Type"]);
        Assert.Equal(source.Content, resource.Content);
    }

    [Fact]
    public async Task SameOriginModeAndRedirectAuthorityRefusalsDoNotReachTheTarget() {
        await using var target = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("private")));
        await using var source = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("", status: 302, headers: "Location: " + target.Origin + "\r\n")));
        var result = await RunAsync("""
            const errors=[];
            try { await fetch(TARGET,{mode:'same-origin'}); } catch(e) { errors.push(e.name); }
            try { await fetch('/redirect'); } catch(e) { errors.push(e.name); }
            return errors;
            """.Replace("TARGET", JsonSerializer.Serialize(target.Origin)), new HtmlScriptRequest { DocumentUrl = source.Origin, ResourcePolicy = new() { AllowNetwork = true } });
        Assert.Equal(2, result.GetArrayLength());
        Assert.Empty(target.Requests);
        Assert.Single(source.Requests);
    }

    [Fact]
    public async Task RedirectsToAnotherAllowedOriginRemoveAuthorizationAndRequireCors() {
        await using var target = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("ok", headers: "Access-Control-Allow-Origin: *\r\n")));
        await using var source = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("", status: 302, headers: "Location: " + target.Origin + "\r\n")));
        var result = await RunAsync("const r=await fetch('/redirect',{headers:{Authorization:'Bearer test'}});return {type:r.type,text:await r.text()};",
            new HtmlScriptRequest { DocumentUrl = source.Origin, ResourcePolicy = new() { AllowNetwork = true, AllowedOrigins = new[] { target.Origin } } });
        Assert.Equal("cors", result.GetProperty("type").GetString());
        var request = Assert.Single(target.Received);
        Assert.False(request.Headers.ContainsKey("Authorization"));
        Assert.Equal(source.Origin.GetLeftPart(UriPartial.Authority), request.Headers["Origin"]);
    }

    [Fact]
    public async Task SameOriginModeRejectsEvenAnOriginPermittedByTheHost() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("unexpected", headers: "Access-Control-Allow-Origin: *\r\n")));
        var result = await RunAsync("try { await fetch(TARGET,{mode:'same-origin'}); } catch(e) { return String(e); } return 'unexpected';".Replace("TARGET", JsonSerializer.Serialize(server.Origin)),
            new HtmlScriptRequest { DocumentUrl = Origin, ResourcePolicy = new() { AllowNetwork = true, AllowedOrigins = new[] { server.Origin } } });
        Assert.Contains("same-origin mode", result.GetString());
        Assert.Empty(server.Requests);
    }

    [Fact]
    public async Task ConcurrentFetchesRespectAdmissionAndShareDeadlines() {
        int active = 0, maximum = 0;
        await using var server = new RuntimeHttpFixture(async (_, token) => {
            int count = Interlocked.Increment(ref active);
            int seen;
            do { seen = maximum; } while (count > seen && Interlocked.CompareExchange(ref maximum, count, seen) != seen);
            try { await Task.Delay(120, token); return RuntimeHttpFixture.Reply.Text("ok"); }
            finally { Interlocked.Decrement(ref active); }
        });
        var result = await RunAsync("return await Promise.all(Array.from({length:5},(_,i)=>fetch('/item'+i).then(r=>r.text())));",
            new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true, MaxConcurrentRequests = 2 } });
        Assert.Equal(5, result.GetArrayLength());
        Assert.InRange(maximum, 1, 2);
        Assert.Equal(5, server.Requests.Count);
    }

    [Fact]
    public async Task ResourceTimeoutAndResponseLimitsRejectFetchWithoutPoisoningTheSession() {
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/slow") await Task.Delay(500, token);
            return RuntimeHttpFixture.Reply.Text(new string('x', path == "/large" ? 129 : 2), chunked: true);
        });
        var result = await RunAsync("""
            const failures=[];
            for(const path of ['/slow','/large']) { try { await fetch(path); } catch(e) { failures.push(String(e)); } }
            const recovery=await (await fetch('/ok')).text();
            return {failures,recovery};
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true, Timeout = TimeSpan.FromMilliseconds(250), MaxResourceBytes = 128 } });
        Assert.Contains("deadline", result.GetProperty("failures")[0].GetString());
        Assert.Contains("byte budget", result.GetProperty("failures")[1].GetString());
        Assert.Equal("xx", result.GetProperty("recovery").GetString());
    }

    [Fact]
    public async Task AbortAfterResponsePreventsUnreadBodyConsumptionAndRepeatedAbortKeepsItsReason() {
        var result = await RunAsync("""
            const controller=new AbortController();
            let events=0; const listener=()=>events++;
            controller.signal.addEventListener('abort',listener);
            controller.signal.addEventListener('abort',listener);
            controller.signal.onabort=()=>events++;
            const r=await fetch('/data',{signal:controller.signal});
            controller.abort('first'); controller.abort('second');
            let reason; try { await r.text(); } catch(e) { reason=e; }
            return {reason,events,signalReason:controller.signal.reason};
            """, new HtmlScriptRequest { DocumentUrl = Origin, Resources = new[] { HtmlRuntimeResource.FromText(new Uri(Origin, "data"), "content", "text/plain") } });
        Assert.Equal("first", result.GetProperty("reason").GetString());
        Assert.Equal("first", result.GetProperty("signalReason").GetString());
        Assert.Equal(2, result.GetProperty("events").GetInt32());
    }

    [Fact]
    public async Task InvalidJsonConsumesTheBodyAndNullBodiesRemainEmpty() {
        var result = await RunAsync("""
            const r=await fetch('/invalid');
            let jsonError; try { await r.json(); } catch(e) { jsonError=e.name; }
            const empty=await fetch('/empty');
            const text=await empty.text();
            return {jsonError,used:r.bodyUsed,text,emptyUsed:empty.bodyUsed,body:empty.body};
            """, new HtmlScriptRequest { DocumentUrl = Origin, Resources = new[] {
                HtmlRuntimeResource.FromText(new Uri(Origin, "invalid"), "broken json", "application/json"),
                new HtmlRuntimeResource(new Uri(Origin, "empty"), Encoding.UTF8.GetBytes("ignored"), "text/plain", 204)
            } });
        Assert.Equal("SyntaxError", result.GetProperty("jsonError").GetString());
        Assert.True(result.GetProperty("used").GetBoolean());
        Assert.Equal("", result.GetProperty("text").GetString());
        Assert.False(result.GetProperty("emptyUsed").GetBoolean());
        Assert.Equal(JsonValueKind.Null, result.GetProperty("body").ValueKind);
    }

    [Fact]
    public async Task UnhandledFetchRejectionFailsCapture() {
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<script>fetch('/missing');</script>", ReadyExpression = "false"
        }));
        Assert.Contains("network loading is disabled", failure.Message);
        Assert.Equal(new[] { new Uri("https://officeimo.invalid/missing") }, failure.MissingResourceUrls);
    }

    [Fact]
    public async Task MissingFetchDiscoveryPreservesQueryAndStripsClientFragment() {
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<script>fetch('/api/report.json?view=summary#client');</script>", ReadyExpression = "false"
        }));

        Assert.Equal(new[] { new Uri("https://officeimo.invalid/api/report.json?view=summary") },
            failure.MissingResourceUrls);
        Assert.Equal(new Uri("https://officeimo.invalid/api/report.json?view=summary"),
            Assert.Single(failure.MissingFetchRequests).Request.Url);
    }

    [Fact]
    public async Task HeaderedMissingFetchIsNotAdvertisedForUrlOnlyReplay() {
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<script>fetch('/vary',{headers:{'X-Variant':'private'}});</script>", ReadyExpression = "false"
        }));

        Assert.Contains("network loading is disabled", failure.Message);
        Assert.Empty(failure.MissingResourceUrls);
    }

    [Fact]
    public async Task LaterFatalMissingGetOutranksAnEarlierHandledMiss() {
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<script>fetch('/optional').catch(() => {}).then(() => fetch('/required'));</script>",
            ReadyExpression = "false"
        }));

        Assert.Contains("network loading is disabled", failure.Message);
        Assert.Contains(new Uri("https://officeimo.invalid/required"), failure.MissingResourceUrls);
    }

    [Fact]
    public async Task ConcurrentFatalAndCaughtMissingGetsRetainBothAttemptedUrls() {
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<script>fetch('/required'); fetch('/optional').catch(() => {});</script>",
            ReadyExpression = "false"
        }));

        Assert.Contains("network loading is disabled", failure.Message);
        Assert.Equal(new[] { "/optional", "/required" },
            failure.MissingResourceUrls.Select(url => url.AbsolutePath).OrderBy(path => path));
    }

    [Theory]
    [InlineData(301)]
    [InlineData(302)]
    [InlineData(303)]
    [InlineData(307)]
    [InlineData(308)]
    public async Task RedirectStatusWithoutLocationReturnsTheOriginalResponse(int status) {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("readable", "text/plain", status)));
        var result = await RunAsync("""
            const r=await fetch('/status');
            const text=await r.text();
            let error; try { await fetch('/error',{redirect:'error'}); } catch(e) { error=e.name; }
            return {status:r.status,redirected:r.redirected,url:r.url,text,error};
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true, MaxRedirects = 0 } });
        Assert.Equal(status, result.GetProperty("status").GetInt32());
        Assert.False(result.GetProperty("redirected").GetBoolean());
        Assert.Equal(new Uri(server.Origin, "status").AbsoluteUri, result.GetProperty("url").GetString());
        Assert.Equal("readable", result.GetProperty("text").GetString());
        Assert.Equal("TypeError", result.GetProperty("error").GetString());
        Assert.Equal(new[] { "/status", "/error" }, server.Requests);
    }

    [Fact]
    public async Task UrlInputsAndUrlSearchParamsBodiesUseTheirBrowserSerialization() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("ok")));
        var result = await RunAsync("""
            const url=new URL('/form?existing=1',document.URL);
            const values=new URLSearchParams(); values.append('name','A B');values.append('name','żółw');
            await fetch(url,{method:'POST',body:values});
            return url.href;
            """, new HtmlScriptRequest { DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true } });
        Assert.Equal(new Uri(server.Origin, "form?existing=1").AbsoluteUri, result.GetString());
        var request = Assert.Single(server.Received);
        Assert.Equal("application/x-www-form-urlencoded;charset=UTF-8", request.Headers["Content-Type"]);
        Assert.Equal("name=A+B&name=%C5%BC%C3%B3%C5%82w", Encoding.UTF8.GetString(request.Body));
    }

    [Fact]
    public async Task UrlSearchParamsPreserveDuplicateOrderAndStayLinkedToTheUrl() {
        var result = await RunAsync("""
            const url=new URL('https://app.example/?name=A+B&name=%C5%BC%C3%B3%C5%82w');
            const params=url.searchParams;
            const initial=params.getAll('name');
            params.append('symbols','+&=!*()~.'); const encoded=url.search;
            url.search='?x=1&x=2&value=a=b=c';
            const same=params===url.searchParams;
            const before=params.getAll('x');
            params.delete('x','1');params.set('value','🐢');params.append('a','second');params.append('a','third');params.sort();
            const copy=new URLSearchParams(params);
            params.set('value','changed');
            const malformed=new URLSearchParams('bad=%E9&literal=%XX&surrogate=\uD800&&');
            return {initial,encoded,same,before,value:copy.get('value'),pairs:Array.from(copy),size:copy.size,url:url.search,malformed:malformed.toString()};
            """);
        Assert.Equal(new[] { "A B", "żółw" }, result.GetProperty("initial").EnumerateArray().Select(value => value.GetString()));
        Assert.EndsWith("&symbols=%2B%26%3D%21*%28%29%7E.", result.GetProperty("encoded").GetString());
        Assert.True(result.GetProperty("same").GetBoolean());
        Assert.Equal(new[] { "1", "2" }, result.GetProperty("before").EnumerateArray().Select(value => value.GetString()));
        Assert.Equal("🐢", result.GetProperty("value").GetString());
        Assert.Equal(4, result.GetProperty("size").GetInt32());
        Assert.Equal("second", result.GetProperty("pairs")[0][1].GetString());
        Assert.Equal("third", result.GetProperty("pairs")[1][1].GetString());
        Assert.Equal("?a=second&a=third&value=changed&x=2", result.GetProperty("url").GetString());
        Assert.Equal("bad=%EF%BF%BD&literal=%25XX&surrogate=%EF%BF%BD", result.GetProperty("malformed").GetString());
    }
}
