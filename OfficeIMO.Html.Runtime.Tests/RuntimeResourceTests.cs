using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Dom;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeResourceTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
    private static readonly Uri Origin = new("https://app.example/");

    [Fact]
    public void ChildFrameLimitsAreValidatedAndRetainedByTheRequestSnapshot() {
        var request = new HtmlScriptRequest {
            MaxChildFrameRealms = 7,
            MaxFrameMessages = 19,
            MaxFrameMessageCharacters = 4096
        };
        HtmlScriptRequest snapshot = request.Snapshot();
        Assert.Equal(7, snapshot.MaxChildFrameRealms);
        Assert.Equal(19, snapshot.MaxFrameMessages);
        Assert.Equal(4096, snapshot.MaxFrameMessageCharacters);
        request.MaxChildFrameRealms = 0;
        Assert.Throws<ArgumentOutOfRangeException>(request.Snapshot);
        request.MaxChildFrameRealms = 1;
        request.MaxFrameMessages = 0;
        Assert.Throws<ArgumentOutOfRangeException>(request.Snapshot);
        request.MaxFrameMessages = 1;
        request.MaxFrameMessageCharacters = 0;
        Assert.Throws<ArgumentOutOfRangeException>(request.Snapshot);
    }

    [Fact]
    public async Task SuppliedExternalScriptsAndStylesheetsUseTheDocumentUrlAndRemainAvailableAfterDisposal() {
        var script = HtmlRuntimeResource.FromText(new Uri(Origin, "scripts/app.js"), "document.querySelector('#status').textContent='Loaded';window.ready=true;", "text/javascript");
        var css = HtmlRuntimeResource.FromText(new Uri(Origin, "styles/app.css"), "#status{color:rgb(0, 128, 0)}", "text/css");
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "reports/index.html"),
            Html = "<link rel='stylesheet' href='../styles/app.css'><p id='status'>Pending</p><script src='../scripts/app.js'></script>",
            Resources = new[] { script, css }
        });
        var capture = await session.CaptureAsync("window.ready===true");
        await session.DisposeAsync();
        Assert.Equal("Loaded", capture.Document.QuerySelector("#status")!.TextContent);
        Assert.Equal(new Uri(Origin, "reports/index.html"), capture.DocumentUrl);
        Assert.Equal(2, capture.Resources.Count);
        var loaded = Assert.Single(capture.Resources, resource => resource.Url == script.Url);
        byte[] mutable = loaded.Content;
        mutable[0] = 0;
        Assert.Equal(script.Content, loaded.Content);
        Assert.Equal(css.Content, Assert.Single(capture.Resources, resource => resource.Url == css.Url).Content);
    }

    [Fact]
    public async Task PictureLoadsTheActiveViewportSourceWithoutRequestingItsFallback() {
        Uri wide = new(Origin, "wide.svg");
        Uri fallback = new(Origin, "fallback.svg");
        var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "report"),
            ViewportWidth = 816D,
            ViewportHeight = 720D,
            Html = """
                <picture>
                  <source media="(max-width:900px)" type="image/svg+xml" srcset="/wide.svg">
                  <source media="(min-width:901px)" type="image/svg+xml" srcset="/narrow.svg">
                  <img src="/fallback.svg" alt="fixture">
                </picture>
                """,
            Resources = new[] { HtmlRuntimeResource.FromText(wide,
                "<svg xmlns='http://www.w3.org/2000/svg' width='2' height='2'></svg>", "image/svg+xml") }
        });

        Assert.Contains(capture.Resources, resource => resource.Url == wide);
        Assert.DoesNotContain(capture.Resources, resource => resource.Url == fallback);
    }

    [Fact]
    public async Task ResponsiveImageLoadsOnlyTheCandidateSelectedForViewportAndDeviceDensity() {
        Uri small = new(Origin, "small.svg");
        Uri medium = new(Origin, "medium.svg");
        Uri large = new(Origin, "large.svg");
        Uri fallback = new(Origin, "fallback.svg");
        string Svg(string color) => $"<svg xmlns='http://www.w3.org/2000/svg' width='2' height='2'><rect width='2' height='2' fill='{color}'/></svg>";
        await using IHtmlRuntimeSession session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri(Origin, "report"),
            ViewportWidth = 800D,
            ViewportHeight = 600D,
            DevicePixelRatio = 2D,
            Html = """
                <img src="/fallback.svg"
                     srcset="/small.svg 400w, /medium.svg 800w, /large.svg 1200w"
                     sizes="(max-width:600px) 100vw, 50vw" alt="fixture">
                """,
            Resources = new[] {
                HtmlRuntimeResource.FromText(small, Svg("red"), "image/svg+xml"),
                HtmlRuntimeResource.FromText(medium, Svg("blue"), "image/svg+xml"),
                HtmlRuntimeResource.FromText(large, Svg("green"), "image/svg+xml"),
                HtmlRuntimeResource.FromText(fallback, Svg("black"), "image/svg+xml")
            }
        });

        Assert.True((await session.EvaluateAsync("devicePixelRatio === 2")).GetBoolean());
        HtmlScriptCapture capture = await session.CaptureAsync();
        Assert.Contains(capture.Resources, resource => resource.Url == medium);
        Assert.DoesNotContain(capture.Resources, resource => resource.Url == small);
        Assert.DoesNotContain(capture.Resources, resource => resource.Url == large);
        Assert.DoesNotContain(capture.Resources, resource => resource.Url == fallback);
    }

    [Fact]
    public async Task SuppliedSameOriginFrameExecutesClassicScriptsInAnIsolatedRealm() {
        Uri frame = new(Origin, "frames/detail.html");
        Uri frameScript = new(Origin, "frames/frame.js");
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "reports/index.html"),
            Html = "<main><p>Outer</p><iframe src='../frames/detail.html'></iframe></main>",
            Resources = new[] {
                HtmlRuntimeResource.FromText(frame, """
                    <!doctype html>
                    <body onload="parent.document.body.dataset.childEvent='ran'">
                    <p id="inside">Frame ready</p>
                    <script>document.querySelector('#inside').textContent='inline ran';parent.document.body.dataset.childInline='ran'</script>
                    <script src="frame.js"></script>
                    """, "text/html; charset=utf-8"),
                HtmlRuntimeResource.FromText(frameScript,
                    "document.querySelector('#inside').textContent='external ran';parent.document.body.dataset.childExternal='ran'",
                    "text/javascript")
            }
        });

        await session.WaitForAsync(
            "document.querySelector('iframe')?.contentDocument?.querySelector('#inside')?.textContent === 'external ran' && document.body.dataset.childEvent === 'ran' && document.body.dataset.childInline === 'ran' && document.body.dataset.childExternal === 'ran'");
        HtmlScriptCapture capture = await session.CaptureAsync();

        Assert.Contains(capture.Resources, resource => resource.Url == frame);
        Assert.Contains(capture.Resources, resource => resource.Url == frameScript);
        Assert.Equal("Outer", capture.Document.QuerySelector("main > p")!.TextContent);
        Assert.Null(capture.Document.QuerySelector("#inside"));
        HtmlElement iframe = capture.Document.QuerySelector("iframe")!;
        Assert.Equal("../frames/detail.html", iframe.GetAttribute("src"));
        HtmlFrameCapture capturedFrame = Assert.Single(capture.Frames);
        Assert.Equal(iframe.NodeId, capturedFrame.FrameElementNodeId);
        Assert.Equal(frame, capturedFrame.DocumentUrl);
        Assert.Equal(frame, capturedFrame.BaseUri);
        Assert.Equal(HtmlDocumentMode.Standards, capturedFrame.Document.Mode);
        Assert.Equal("external ran", capturedFrame.Document.QuerySelector("#inside")!.TextContent);
        Assert.Empty(capturedFrame.Frames);
        HtmlRuntimeArtifactEntry frameEntry = Assert.Single(capture.ArtifactManifest.Entries,
            entry => entry.Name == "frame-0001/document.html");
        Assert.Equal(System.Text.Encoding.UTF8.GetByteCount(capturedFrame.Document.OuterHtml), frameEntry.ByteCount);
        string json = HtmlRuntimeJson.Serialize(capture);
        Assert.Contains("external ran", json, StringComparison.Ordinal);
        using (System.Text.Json.JsonDocument payload = System.Text.Json.JsonDocument.Parse(json)) {
            System.Text.Json.JsonElement framePayload = payload.RootElement.GetProperty("frames")[0];
            Assert.Contains("<!DOCTYPE html>", framePayload.GetProperty("documentHtml").GetString(), StringComparison.OrdinalIgnoreCase);
            Assert.Equal("Standards", framePayload.GetProperty("documentMode").GetString());
        }

        HtmlDocument renderDocument = capture.CreateRenderDocument();
        Assert.Null(capture.Document.QuerySelector("iframe")!.GetAttribute("srcdoc"));
        string srcdoc = renderDocument.QuerySelector("iframe")!.GetAttribute("srcdoc")!;
        Assert.Contains("<!DOCTYPE html>", srcdoc, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("external ran", srcdoc, StringComparison.Ordinal);
    }

    [Fact]
    public async Task NestedSameOriginFramesRemainASeparateBoundedCaptureTree() {
        Uri firstUrl = new(Origin, "frames/first.html");
        Uri secondUrl = new(Origin, "frames/nested/second.html");
        HtmlScriptCapture capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "index.html"),
            Html = "<iframe src='/frames/first.html'></iframe>",
            Resources = new[] {
                HtmlRuntimeResource.FromText(firstUrl,
                    "<body><p id='first'>First frame</p><script>document.body.dataset.realm=parent!==window&&top!==window</script><iframe src='nested/second.html'></iframe></body>",
                    "text/html; charset=utf-8"),
                HtmlRuntimeResource.FromText(secondUrl,
                    "<body><p id='second'>Second frame</p><script>document.body.dataset.realm=parent!==window&&top!==parent;parent.document.body.dataset.grandchild='ran';Promise.resolve().then(()=>document.body.dataset.microtask='ran');setTimeout(()=>document.body.dataset.timer='ran',1)</script></body>",
                    "text/html; charset=utf-8")
            },
            ReadyExpression = "document.querySelector('iframe')?.contentDocument?.querySelector('iframe')?.contentDocument?.body?.dataset.timer==='ran'"
        });

        HtmlFrameCapture first = Assert.Single(capture.Frames);
        HtmlFrameCapture second = Assert.Single(first.Frames);
        Assert.Equal("First frame", first.Document.QuerySelector("#first")!.TextContent);
        Assert.Equal("Second frame", second.Document.QuerySelector("#second")!.TextContent);
        Assert.Equal("true", first.Document.Body!.GetAttribute("data-realm"));
        Assert.Equal("ran", first.Document.Body.GetAttribute("data-grandchild"));
        Assert.Equal("true", second.Document.Body!.GetAttribute("data-realm"));
        Assert.Equal("ran", second.Document.Body.GetAttribute("data-microtask"));
        Assert.Equal("ran", second.Document.Body.GetAttribute("data-timer"));
        Assert.Contains("Second frame", capture.CreateRenderDocument().QuerySelector("iframe")!.GetAttribute("srcdoc"), StringComparison.Ordinal);
        Assert.Equal(3, capture.ArtifactManifest.Entries.Count(entry => entry.Name.EndsWith("document.html", StringComparison.Ordinal)));
    }

    [Fact]
    public async Task SameOriginFrameGlobalsAndMessagingRemainRealmScoped() {
        Uri frame = new(Origin, "messaging-frame.html");
        await using IHtmlRuntimeSession session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "index.html"),
            Html = """
                <body><script>
                addEventListener('message',event=>{
                  if(event.data.kind!=='ready')return;
                  document.body.dataset.received=event.data.value;
                  document.body.dataset.origin=event.origin;
                  document.body.dataset.source=event.source===document.querySelector('#child').contentWindow;
                  event.source.postMessage({kind:'reply',value:42},event.origin);
                });
                </script><iframe id="child" src="/messaging-frame.html"></iframe></body>
                """,
            Resources = new[] {
                HtmlRuntimeResource.FromText(frame, """
                    <body><output id="result">waiting</output>
                    <script>
                    window.childOnly='child';
                    document.body.dataset.identities=[window===self,parent!==window,top===parent,frameElement?.id==='child'].join(',');
                    parent.document.body.dataset.childAccess='yes';
                    parent.document.body.dataset.postType=typeof parent.postMessage;
                    try{parent.postMessage(()=>{},'*')}catch(e){document.body.dataset.cloneError=e.name}
                    try{parent.postMessage({},'*',[new ArrayBuffer(1)])}catch(e){document.body.dataset.transferError=e.name}
                    try{parent.postMessage({},'invalid-origin')}catch(e){document.body.dataset.originError=e.name}
                    addEventListener('message',event=>{
                      document.querySelector('#result').textContent=event.data.kind+':'+event.data.value;
                      document.body.dataset.source=event.source===parent;
                    });
                    try{parent.postMessage({kind:'ready',value:'child-ready'},'*')}catch(e){parent.document.body.dataset.postError=e.name+':'+e.message}
                    </script></body>
                    """, "text/html; charset=utf-8")
            }
        });

        await session.WaitForAsync("document.body.dataset.childAccess==='yes'");
        var state = await session.EvaluateAsync("({result:document.querySelector('#child').contentDocument.querySelector('#result').textContent,received:document.body.dataset.received??null,origin:document.body.dataset.origin??null,source:document.body.dataset.source??null,access:document.body.dataset.childAccess,postType:document.body.dataset.postType,postError:document.body.dataset.postError??null,cloneError:document.querySelector('#child').contentDocument.body.dataset.cloneError,transferError:document.querySelector('#child').contentDocument.body.dataset.transferError,originError:document.querySelector('#child').contentDocument.body.dataset.originError,isolated:typeof document.querySelector('#child').contentWindow.childOnly==='undefined',identities:document.querySelector('#child').contentDocument.body.dataset.identities,childSource:document.querySelector('#child').contentDocument.body.dataset.source??null})");

        Assert.Equal("function", state.GetProperty("postType").GetString());
        Assert.Equal(System.Text.Json.JsonValueKind.Null, state.GetProperty("postError").ValueKind);
        Assert.Equal("DataCloneError", state.GetProperty("cloneError").GetString());
        Assert.Equal("DataCloneError", state.GetProperty("transferError").GetString());
        Assert.Equal("SyntaxError", state.GetProperty("originError").GetString());
        Assert.Equal("child-ready", state.GetProperty("received").GetString());
        Assert.Equal("https://app.example", state.GetProperty("origin").GetString());
        Assert.Equal("true", state.GetProperty("source").GetString());
        Assert.Equal("yes", state.GetProperty("access").GetString());
        Assert.True(state.GetProperty("isolated").GetBoolean());
        Assert.Equal("true,true,true,true", state.GetProperty("identities").GetString());
        Assert.Equal("true", state.GetProperty("childSource").GetString());
        Assert.Equal("reply:42", state.GetProperty("result").GetString());
    }

    [Fact]
    public async Task FrameMessagingUsesIndependentStructuredCloneGraphs() {
        Uri frame = new(Origin, "structured-message-frame.html");
        await using IHtmlRuntimeSession session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "index.html"),
            Html = """
                <body><script>
                addEventListener('message',event=>{
                  const data=event.data, key=[...data.map.keys()][0];
                  window.cloneResult={cycle:data.self===data,alias:data.shared===data.again,mapAlias:key===data.shared,
                    set:data.set.has(data.shared),date:data.date instanceof Date&&data.date.toISOString()==='2024-01-02T03:04:05.000Z',
                    regexp:data.regexp instanceof RegExp&&data.regexp.source==='a+b'&&data.regexp.flags.includes('i'),
                    bigint:data.bigint===9007199254740993n,nan:Number.isNaN(data.nan),negativeZero:Object.is(data.negativeZero,-0),
                    infinity:data.infinity===Infinity,typed:data.typed instanceof Uint16Array&&data.typed[0]===513&&data.bytes[0]===1,
                    error:data.error instanceof TypeError&&data.error.message==='bad'&&data.error.cause.code===42,
                    undefinedValue:'missing' in data&&data.missing===undefined,snapshot:data.shared.name==='before'};
                });
                </script><iframe src="/structured-message-frame.html"></iframe></body>
                """,
            Resources = new[] { HtmlRuntimeResource.FromText(frame, """
                <body><script>
                const shared={name:'before'},buffer=new ArrayBuffer(4),bytes=new Uint8Array(buffer);bytes.set([1,2,3,4]);
                const error=new TypeError('bad',{cause:{code:42}});
                const payload={shared,again:shared,map:new Map([[shared,'value']]),set:new Set([shared]),
                  date:new Date('2024-01-02T03:04:05Z'),regexp:/a+b/gi,bigint:9007199254740993n,nan:NaN,
                  negativeZero:-0,infinity:Infinity,typed:new Uint16Array(buffer),bytes,error,missing:undefined};
                payload.self=payload;parent.postMessage(payload,'*');shared.name='after';bytes[0]=9;
                </script></body>
                """, "text/html; charset=utf-8") }
        });

        await session.WaitForAsync("!!window.cloneResult");
        var result = await session.EvaluateAsync("window.cloneResult");
        foreach (System.Text.Json.JsonProperty property in result.EnumerateObject())
            Assert.True(property.Value.GetBoolean(), property.Name);
    }

    [Fact]
    public async Task FrameExecutionRespectsOriginAndSandboxFlags() {
        Uri sandboxed = new(Origin, "sandboxed-script.html");
        Uri allowed = new(Origin, "allowed-script.html");
        Uri externalOrigin = new("https://frames.example/");
        Uri external = new(externalOrigin, "external-script.html");
        await using IHtmlRuntimeSession session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "index.html"),
            Html = "<body><iframe id='blocked' sandbox src='/sandboxed-script.html'></iframe><iframe id='allowed' sandbox='allow-scripts allow-same-origin' src='/allowed-script.html'></iframe><iframe id='external' src='https://frames.example/external-script.html'></iframe></body>",
            ResourcePolicy = new HtmlRuntimeResourcePolicy { AllowedOrigins = new[] { externalOrigin } },
            Resources = new[] {
                HtmlRuntimeResource.FromText(sandboxed, "<p>blocked</p><script>parent.document.body.dataset.blocked='ran'</script>", "text/html; charset=utf-8"),
                HtmlRuntimeResource.FromText(allowed, "<p>allowed</p><script>parent.document.body.dataset.allowed='ran'</script>", "text/html; charset=utf-8"),
                HtmlRuntimeResource.FromText(external, "<p>external</p><script>document.body.dataset.executed='yes'</script>", "text/html; charset=utf-8")
            }
        });

        await session.WaitForAsync("document.querySelector('#external')?.contentDocument?.querySelector('p')?.textContent==='external' && document.body.dataset.allowed==='ran'");
        var state = await session.EvaluateAsync("({blocked:document.body.dataset.blocked??null,allowed:document.body.dataset.allowed,external:document.querySelector('#external').contentDocument.body.dataset.executed??null})");
        Assert.Equal(System.Text.Json.JsonValueKind.Null, state.GetProperty("blocked").ValueKind);
        Assert.Equal("ran", state.GetProperty("allowed").GetString());
        Assert.Equal(System.Text.Json.JsonValueKind.Null, state.GetProperty("external").ValueKind);
    }

    [Theory]
    [InlineData("inline")]
    [InlineData("external")]
    [InlineData("timer")]
    public async Task ChildFrameFailuresReachTheSharedSessionError(string stage) {
        Uri frame = new(Origin, "failing-frame.html");
        Uri script = new(Origin, "failing-frame.js");
        string source = stage switch {
            "external" => "<script src='/failing-frame.js'></script>",
            "timer" => "<script>setTimeout(()=>{throw new Error('child timer failure')},1)</script>",
            _ => "<script>throw new Error('child inline failure')</script>"
        };
        var resources = new List<HtmlRuntimeResource> {
            HtmlRuntimeResource.FromText(frame, source, "text/html; charset=utf-8")
        };
        if (stage == "external") resources.Add(HtmlRuntimeResource.FromText(script,
            "throw new Error('child external failure')", "text/javascript"));

        async Task ExerciseAsync() {
            await using IHtmlRuntimeSession session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
                DocumentUrl = new Uri(Origin, "index.html"),
                Html = "<iframe src='/failing-frame.html'></iframe>",
                Resources = resources.ToArray(),
                Timeout = TimeSpan.FromSeconds(10)
            });
            await session.WaitForAsync("false");
        }

        HtmlScriptRuntimeException failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(ExerciseAsync);
        Assert.Contains("child", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DetachedAndReplacedFramesRetireTheirOriginalRealms() {
        Uri first = new(Origin, "retired-frame.html");
        Uri replacement = new(Origin, "replacement-frame.html");
        await using IHtmlRuntimeSession session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "index.html"),
            Html = "<body><iframe id='child' src='/retired-frame.html'></iframe></body>",
            Resources = new[] {
                HtmlRuntimeResource.FromText(first,
                    "<body><script>parent.document.body.dataset.originalReady='yes';setTimeout(()=>parent.document.body.dataset.stale='ran',500)</script></body>",
                    "text/html; charset=utf-8"),
                HtmlRuntimeResource.FromText(replacement,
                    "<body><p>replacement</p><script>parent.document.body.dataset.replacementScript='ran'</script></body>",
                    "text/html; charset=utf-8")
            }
        });

        await session.WaitForAsync("document.body.dataset.originalReady==='yes'");
        await session.ExecuteAsync("document.querySelector('#child').src='/replacement-frame.html'");
        await session.WaitForAsync("document.querySelector('#child')?.contentDocument?.readyState==='complete'&&document.querySelector('#child').contentDocument.querySelector('p')?.textContent==='replacement'");
        await Task.Delay(750);
        Assert.True((await session.EvaluateAsync(
            "!document.body.dataset.stale&&!document.body.dataset.replacementScript")).GetBoolean());

        await session.ExecuteAsync("document.querySelector('#child').remove()");
        await Task.Delay(25);
        Assert.True((await session.EvaluateAsync("!document.body.dataset.stale")).GetBoolean());
    }

    [Fact]
    public async Task SameOriginRedirectedFrameKeepsItsQualifiedRealm() {
        Uri requested = new(Origin, "redirect-frame.html");
        Uri final = new(Origin, "frames/final-frame.html");
        byte[] html = System.Text.Encoding.UTF8.GetBytes(
            "<body><p>redirected</p><script>document.body.dataset.executed='yes';parent.document.body.dataset.redirected='yes'</script></body>");
        await using IHtmlRuntimeSession session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "index.html"),
            Html = "<body><iframe src='/redirect-frame.html'></iframe></body>",
            Resources = new[] {
                new HtmlRuntimeResource(requested, html, "text/html; charset=utf-8", finalUrl: final, redirectCount: 1)
            }
        });

        await session.WaitForAsync("document.body.dataset.redirected==='yes'");
        var state = await session.EvaluateAsync("({url:document.querySelector('iframe').contentDocument.URL,executed:document.querySelector('iframe').contentDocument.body.dataset.executed})");

        Assert.Equal(final.AbsoluteUri, state.GetProperty("url").GetString());
        Assert.Equal("yes", state.GetProperty("executed").GetString());
    }

    [Fact]
    public async Task MessageQueuedByAFrameIsDiscardedWhenThatFrameDetaches() {
        Uri frame = new(Origin, "self-removing-frame.html");
        await using IHtmlRuntimeSession session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "index.html"),
            Html = "<body><script>onmessage=()=>document.body.dataset.staleMessage='delivered'</script><iframe src='/self-removing-frame.html'></iframe></body>",
            Resources = new[] { HtmlRuntimeResource.FromText(frame,
                "<script>parent.postMessage('stale','*');frameElement.remove()</script>",
                "text/html; charset=utf-8") }
        });

        await Task.Delay(50);
        Assert.True((await session.EvaluateAsync("!document.body.dataset.staleMessage&&!document.querySelector('iframe')")).GetBoolean());
    }

    [Fact]
    public async Task ChildFrameRealmAndMessageBudgetsAreCumulative() {
        Uri first = new(Origin, "realm-one.html");
        Uri second = new(Origin, "realm-two.html");
        HtmlScriptCapture realmCapture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
                DocumentUrl = new Uri(Origin, "index.html"),
                Html = "<iframe src='/realm-one.html'></iframe><iframe src='/realm-two.html'></iframe>",
                Resources = new[] {
                    HtmlRuntimeResource.FromText(first, "<body><script>document.body.dataset.ready='one'</script></body>", "text/html; charset=utf-8"),
                    HtmlRuntimeResource.FromText(second, "<body><script>document.body.dataset.ready='two'</script></body>", "text/html; charset=utf-8")
                },
                MaxChildFrameRealms = 1,
                ReadyExpression = "!!document.querySelectorAll('iframe')[0]?.contentDocument?.body && !!document.querySelectorAll('iframe')[1]?.contentDocument?.body"
            });
        Assert.Equal(new string?[] { "one", null }, realmCapture.Frames.Select(frame => frame.Document.Body!.GetAttribute("data-ready")));

        Uri frame = new(Origin, "message-budget.html");
        HtmlScriptCapture messageCapture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
                DocumentUrl = new Uri(Origin, "index.html"),
                Html = "<iframe src='/message-budget.html'></iframe>",
                Resources = new[] { HtmlRuntimeResource.FromText(frame, "<body><script>parent.postMessage(1,'*');try{parent.postMessage(2,'*')}catch(e){document.body.dataset.error=e.name}document.body.dataset.done='yes'</script></body>", "text/html; charset=utf-8") },
                MaxFrameMessages = 1,
                ReadyExpression = "document.querySelector('iframe')?.contentDocument?.body?.dataset.done==='yes'"
            });
        Assert.Equal("QuotaExceededError", Assert.Single(messageCapture.Frames).Document.Body!.GetAttribute("data-error"));

        Uri largeMessageFrame = new(Origin, "large-message.html");
        HtmlScriptCapture largeMessage = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "index.html"),
            Html = "<iframe src='/large-message.html'></iframe>",
            Resources = new[] { HtmlRuntimeResource.FromText(largeMessageFrame,
                "<body><script>try{parent.postMessage({value:'message-too-large'},'*')}catch(e){document.body.dataset.error=e.name}</script></body>",
                "text/html; charset=utf-8") },
            MaxFrameMessageCharacters = 8,
            ReadyExpression = "document.querySelector('iframe')?.contentDocument?.body?.dataset.error==='QuotaExceededError'"
        });
        Assert.Equal("QuotaExceededError", Assert.Single(largeMessage.Frames).Document.Body!.GetAttribute("data-error"));
    }

    [Fact]
    public async Task CapturedFramesFollowContainingDocumentPreorder() {
        Uri firstUrl = new(Origin, "first.html");
        Uri secondUrl = new(Origin, "second.html");
        HtmlScriptCapture capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "index.html"),
            Html = "<div><iframe id='first' src='/first.html'></iframe></div><iframe id='second' src='/second.html'></iframe>",
            Resources = new[] {
                HtmlRuntimeResource.FromText(firstUrl, "<p>First</p>", "text/html; charset=utf-8"),
                HtmlRuntimeResource.FromText(secondUrl, "<p>Second</p>", "text/html; charset=utf-8")
            },
            ReadyExpression = "document.querySelector('#first')?.contentDocument?.querySelector('p')?.textContent==='First' && document.querySelector('#second')?.contentDocument?.querySelector('p')?.textContent==='Second'"
        });

        Assert.Equal(new[] {
            capture.Document.QuerySelector("#first")!.NodeId,
            capture.Document.QuerySelector("#second")!.NodeId
        }, capture.Frames.Select(frame => frame.FrameElementNodeId));
        Assert.Equal(new[] { firstUrl, secondUrl }, capture.Frames.Select(frame => frame.DocumentUrl));
    }

    [Fact]
    public async Task SandboxedAndCrossOriginFrameBodiesAreNotExposedByCapture() {
        Uri sameOriginFrame = new(Origin, "sandboxed.html");
        Uri externalOrigin = new("https://frames.example/");
        Uri externalFrame = new(externalOrigin, "external.html");
        HtmlScriptCapture capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "index.html"),
            Html = "<iframe sandbox src='/sandboxed.html'></iframe><iframe src='https://frames.example/external.html'></iframe>",
            ResourcePolicy = new HtmlRuntimeResourcePolicy { AllowedOrigins = new[] { externalOrigin } },
            Resources = new[] {
                HtmlRuntimeResource.FromText(sameOriginFrame, "<p>Sandboxed</p>", "text/html; charset=utf-8"),
                HtmlRuntimeResource.FromText(externalFrame, "<p>External</p>", "text/html; charset=utf-8")
            }
        });

        Assert.Empty(capture.Frames);
        Assert.DoesNotContain("Sandboxed", capture.CreateRenderDocument().DocumentElement!.OuterHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("External", capture.CreateRenderDocument().DocumentElement!.OuterHtml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task RemovingSandboxAfterNavigationDoesNotExposeTheOpaqueFrame() {
        Uri frameUrl = new(Origin, "sandboxed-mutation.html");
        await using IHtmlRuntimeSession session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri(Origin, "index.html"),
            Html = "<iframe sandbox src='/sandboxed-mutation.html'></iframe>",
            Resources = new[] {
                HtmlRuntimeResource.FromText(frameUrl, "<p>Opaque child</p>", "text/html; charset=utf-8")
            }
        });
        await session.WaitForAsync(
            "document.querySelector('iframe')?.contentDocument?.querySelector('p')?.textContent==='Opaque child'");

        await session.EvaluateAsync("document.querySelector('iframe').removeAttribute('sandbox')");
        HtmlScriptCapture capture = await session.CaptureAsync();

        Assert.Empty(capture.Frames);
        Assert.DoesNotContain("Opaque child", capture.CreateRenderDocument().OuterHtml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task FrameDocumentsShareTheRootCaptureNodeBudget() {
        Uri frame = new(Origin, "large-frame.html");
        string body = string.Concat(Enumerable.Range(0, 24).Select(index => $"<p>{index}</p>"));
        HtmlScriptRuntimeException error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
                DocumentUrl = new Uri(Origin, "index.html"),
                Html = "<iframe src='/large-frame.html'></iframe>",
                Resources = new[] { HtmlRuntimeResource.FromText(frame, body, "text/html; charset=utf-8") },
                MaxNodes = 20
            }));

        Assert.Contains("node budget", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task FrameCaptureRejectsNestingBeyondItsSupportedDepth() {
        var resources = new List<HtmlRuntimeResource>();
        for (int index = 0; index < 9; index++) {
            string html = index == 8
                ? "<p>Deepest frame</p>"
                : $"<iframe src='/frame-{index + 1}.html'></iframe>";
            resources.Add(HtmlRuntimeResource.FromText(
                new Uri(Origin, $"frame-{index}.html"), html, "text/html; charset=utf-8"));
        }

        HtmlScriptRuntimeException error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
                DocumentUrl = new Uri(Origin, "index.html"),
                Html = "<iframe src='/frame-0.html'></iframe>",
                Resources = resources
            }));

        Assert.Contains("frame depth budget", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task NetworkIsOptInAndAllowedOriginsAreCheckedBeforeRequests() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("window.ready=true")));
        var disabled = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = "<script src='/app.js'></script>"
        }));
        Assert.Contains("network loading is disabled", disabled.Message);
        Assert.Equal(new[] { new Uri(server.Origin, "app.js") }, disabled.MissingResourceUrls);
        var forbidden = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = Origin, Html = $"<script src='{new Uri(server.Origin, "app.js")}'></script>", ResourcePolicy = new() { AllowNetwork = true }
        }));
        Assert.Contains("origin is not allowed", forbidden.Message);
        Assert.Empty(forbidden.MissingResourceUrls);
        Uri approvedExternal = new("https://assets.example/app.js");
        var approvedButUnsupplied = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
                DocumentUrl = Origin, Html = $"<script src='{approvedExternal}'></script>",
                ResourcePolicy = new() { AllowedOrigins = new[] { new Uri("https://assets.example/") } }
            }));
        Assert.Equal(new[] { approvedExternal }, approvedButUnsupplied.MissingResourceUrls);
        Assert.Empty(server.Requests);
        var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = "<script src='/app.js'></script>", ResourcePolicy = new() { AllowNetwork = true }, ReadyExpression = "window.ready===true"
        });
        Assert.Single(server.Requests);
        Assert.Equal(new Uri(server.Origin, "app.js"), Assert.Single(capture.Resources).Url);
    }

    [Fact]
    public async Task RedirectsRecheckOriginsAndRetainFinalIdentity() {
        await using var target = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("window.ready=true")));
        await using var source = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("", status: 302, headers: "Location: " + new Uri(target.Origin, "final.js") + "\r\n")));
        var request = new HtmlScriptRequest { DocumentUrl = source.Origin, Html = "<script src='/start.js'></script>", ResourcePolicy = new() { AllowNetwork = true }, ReadyExpression = "window.ready===true" };
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(request));
        Assert.Contains("origin is not allowed", failure.Message);
        Assert.Empty(target.Requests);
        request.ResourcePolicy.AllowedOrigins = new[] { target.Origin };
        var result = Assert.Single((await Runtime().CaptureTrustedAsync(request)).Resources);
        Assert.Equal(new Uri(source.Origin, "start.js"), result.Url);
        Assert.Equal(new Uri(target.Origin, "final.js"), result.FinalUrl);
        Assert.Equal(1, result.RedirectCount);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ResponseLimitsApplyToDeclaredAndStreamedBytes(bool chunked) {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text(new string(' ', 129), chunked: chunked)));
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = "<script src='/large.js'></script>", ResourcePolicy = new() { AllowNetwork = true, MaxResourceBytes = 128 }
        }));
        Assert.Contains("byte budget", failure.Message);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task RequestAndTotalByteBudgetsSpanTheLiveSession(bool requests) {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("window.loaded=true;")));
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = "<script src='/first.js'></script>",
            ResourcePolicy = new() { AllowNetwork = true, MaxRequests = requests ? 1 : 10, MaxTotalBytes = requests ? 4096 : 20 }
        });
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(async () => {
            await session.ExecuteAsync("const script=document.createElement('script');script.src='/second.js';document.head.appendChild(script);");
            await session.CaptureAsync("false");
        });
        Assert.Contains(requests ? "request budget" : "byte budget", failure.Message);
    }

    [Fact]
    public async Task CallerCancellationInterruptsAResourceLoad() {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (_, token) => {
            started.TrySetResult();
            await Task.Delay(System.Threading.Timeout.Infinite, token);
            return RuntimeHttpFixture.Reply.Text("");
        });
        using var cancel = new CancellationTokenSource();
        Task capture = Runtime().CaptureTrustedAsync(new HtmlScriptRequest { DocumentUrl = server.Origin, Html = "<script src='/slow.js'></script>", ResourcePolicy = new() { AllowNetwork = true } }, cancel.Token);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancel.Cancel();
        var failure = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => capture);
        Assert.Equal(cancel.Token, failure.CancellationToken);
    }

    [Theory]
    [InlineData("file:///tmp/private.js")]
    [InlineData("data:text/javascript,window.ready=true")]
    public async Task OtherSchemesDoNotBypassTheRequester(string url) {
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest { Html = $"<script src='{url}'></script>", ResourcePolicy = new() { AllowNetwork = true } }));
        Assert.Contains("HTTP(S)", failure.Message);
    }

    [Fact]
    public async Task ResourceDeadlineAndHttpFailurePreventSuccessfulCaptures() {
        await using var slow = new RuntimeHttpFixture(async (_, token) => { await Task.Delay(1000, token); return RuntimeHttpFixture.Reply.Text(""); });
        var timeout = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = slow.Origin, Html = "<script src='/slow.js'></script>", ResourcePolicy = new() { AllowNetwork = true, Timeout = TimeSpan.FromMilliseconds(100) }
        }));
        Assert.Contains("resource load exceeded its deadline", timeout.Message);
        await using var missing = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("not found", status: 404)));
        var failure = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = missing.Origin, Html = "<script src='/missing.js'></script>", ResourcePolicy = new() { AllowNetwork = true }
        }));
        Assert.Contains("HTTP 404", failure.Message);
    }

    [Fact]
    public async Task RedirectLoopsAndUrlCredentialsAreRejected() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("", status: 302, headers: "Location: /again.js\r\n")));
        var loop = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = "<script src='/again.js'></script>", ResourcePolicy = new() { AllowNetwork = true, MaxRedirects = 1 }
        }));
        Assert.Contains("redirect budget", loop.Message);
        Assert.Equal(2, server.Requests.Count);
        var credentialUrl = new UriBuilder(server.Origin) { UserName = "test", Password = "test", Path = "/credentials.js" }.Uri;
        var credentials = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = server.Origin, Html = $"<script src='{credentialUrl}'></script>", ResourcePolicy = new() { AllowNetwork = true }
        }));
        Assert.Contains("without credentials", credentials.Message);
        Assert.Equal(2, server.Requests.Count);
    }

    [Fact]
    public async Task SuppliedResourceIdentityAndBudgetsAreValidatedBeforeOpening() {
        var upper = HtmlRuntimeResource.FromText(new Uri(Origin, "A.js"), "window.upper=true", "text/javascript");
        var lower = HtmlRuntimeResource.FromText(new Uri(Origin, "a.js"), "window.lower=true", "text/javascript");
        var request = new HtmlScriptRequest { DocumentUrl = Origin, Html = "<script src='/A.js'></script><script src='/a.js'></script>", Resources = new[] { upper, lower }, ReadyExpression = "window.upper&&window.lower" };
        Assert.Equal(2, (await Runtime().CaptureTrustedAsync(request)).Resources.Count);
        request.Resources = new[] { upper, upper };
        await Assert.ThrowsAsync<ArgumentException>(() => Runtime().OpenTrustedAsync(request));
        request.Resources = new[] { upper };
        request.ResourcePolicy.MaxResourceBytes = 1;
        await Assert.ThrowsAsync<ArgumentException>(() => Runtime().OpenTrustedAsync(request));
    }
}
