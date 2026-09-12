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
