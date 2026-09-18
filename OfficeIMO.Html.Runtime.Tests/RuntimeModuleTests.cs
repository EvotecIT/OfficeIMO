using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeModuleTests {
    private static readonly Uri Page = new("https://modules.example/nested/page.html");
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
    private static HtmlRuntimeResource Source(string path, string source, string type = "text/javascript") => HtmlRuntimeResource.FromText(new Uri(Page, path), source, type);
    private static string Integrity(string source) => "sha256-" + Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(source)));

    [Fact]
    public async Task ModuleApplicationAwaitsFetchAndTimersThenUsesDynamicImportsAndIndependentCapture() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = "<h1>Module report</h1><p id='total'>Loading</p><button>Update</button><script type='module' src='/app/main.js'></script>",
            Resources = new[] {
                Source("/app/main.js", """
                    import * as counter from './counter.js';
                    import {read} from './reader.js';
                    const data=await fetch('/data.json').then(r=>r.json());
                    await new Promise(resolve=>setTimeout(resolve,0));
                    window.identity=counter===await import('/app/counter.js');
                    window.moduleUrl=import.meta.url;
                    window.liveBefore=read();
                    document.querySelector('#total').textContent='Total: '+data.total;
                    document.querySelector('button').onclick=async()=>{
                        const format=await import('./format.js');
                        counter.increment();
                        document.querySelector('#total').textContent=format.label(data.total+read());
                    };
                    """),
                Source("/app/counter.js", "export let value=1;export function increment(){value++}"),
                Source("/app/reader.js", "import {value} from './counter.js';export function read(){return value}"),
                Source("/app/format.js", "export function label(value){return 'Total: '+value}"),
                Source("/data.json", "{\"total\":40}", "application/json")
            }
        });
        Assert.True((await session.EvaluateAsync("identity && liveBefore===1")).GetBoolean());
        Assert.Equal("https://modules.example/app/main.js", (await session.EvaluateAsync("moduleUrl")).GetString());
        var before = await session.CaptureAsync();
        await session.Locator("button").ClickAsync();
        await session.Locator("#total").WaitForTextAsync("Total: 42");
        var after = await session.CaptureAsync();
        await session.DisposeAsync();
        Assert.Equal("Total: 40", before.Document.QuerySelector("#total")!.TextContent);
        Assert.Equal("Total: 42", after.Document.QuerySelector("#total")!.TextContent);
        Assert.Contains(after.Resources, item => item.Url.AbsolutePath == "/app/format.js");
        var conversion = HtmlConversionDocument.FromDocument(after.Document);
        Assert.Contains("Total: 42", conversion.ToMarkdown());
        Assert.Contains("Total: 42", PdfReadDocument.Open(conversion.ToPdfBytes()).ExtractText());
    }

    [Fact]
    public async Task InlineModulesUseTheDocumentBaseAndPreserveDistinctBindings() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = "<base href='/assets/'><script type='module'>import {value} from './dep.js';const local=1;window.first=value+local;window.url=import.meta.url</script><script type='module'>const local=2;window.second=local</script>",
            Resources = new[] { Source("/assets/dep.js", "export const value=40") }
        });
        Assert.Equal(41, (await session.EvaluateAsync("first")).GetInt32());
        Assert.Equal(2, (await session.EvaluateAsync("second")).GetInt32());
        Assert.Equal("https://modules.example/assets/", (await session.EvaluateAsync("url")).GetString());
        Assert.Equal("undefined", (await session.EvaluateAsync("typeof local")).GetString());
    }

    [Fact]
    public async Task ImportMapScopesAndPrefixesUseNormalizedUrlsAndLongestMatchingScope() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = """
                <script type='importmap'>{"imports":{"dep":"/global.js","lib/":"/library/"},"scopes":{"/app/":{"dep":"/scoped.js"},"/app/deep/":{"dep":"/deep.js"}}}</script>
                <script type='module' src='/app/deep/main.js'></script>
                """,
            Resources = new[] {
                Source("/app/deep/main.js", "import {value} from 'dep';import {other} from 'lib/other.js';window.result=value+other"),
                Source("/global.js", "export const value=1"), Source("/scoped.js", "export const value=2"), Source("/deep.js", "export const value=40"),
                Source("/library/other.js", "export const other=2")
            }
        });
        Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
    }

    [Fact]
    public async Task ChildFrameModulesUseTheirOwnRealmImportMapAndJsonAttributes() {
        Uri frame = new(Page, "/frames/report.html");
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = """
                <script>window.childOnly='root';addEventListener('message',event=>{if(event.data?.kind==='frame-module')window.frameResult=event.data})</script>
                <script type='importmap'>{"imports":{"dep":"/wrong.js"}}</script>
                <iframe src='/frames/report.html'></iframe>
                """,
            Resources = new[] {
                HtmlRuntimeResource.FromText(frame, """
                    <script type='importmap'>{"imports":{"dep":"./dep.js","settings":"./settings.json"}}</script>
                    <p id='result'>Loading child modules</p>
                    <script type='module'>
                        import {label} from 'dep';
                        import settings from 'settings' with {type:'json'};
                        const extra=await import('./extra.json',{with:{type:'json'}});
                        globalThis.childOnly=`${label} ${settings.value+extra.default.delta}`;
                        document.querySelector('#result').textContent=childOnly;
                        parent.postMessage({kind:'frame-module',value:childOnly,url:import.meta.url},'*');
                    </script>
                    """, "text/html; charset=utf-8"),
                Source("/frames/dep.js", "export const label='Frame modules ready'"),
                Source("/frames/settings.json", "{\"value\":40}", "application/json; charset=utf-8"),
                Source("/frames/extra.json", "{\"delta\":2}", "application/vnd.officeimo+json"),
                Source("/wrong.js", "throw new Error('The root import map leaked into the child frame')")
            }
        });

        await session.WaitForAsync("window.frameResult?.value==='Frame modules ready 42'");

        Assert.True((await session.EvaluateAsync("childOnly==='root'")).GetBoolean());
        Assert.Equal(frame.AbsoluteUri, (await session.EvaluateAsync("frameResult.url")).GetString());
        HtmlScriptCapture capture = await session.CaptureAsync();
        Assert.Equal("Frame modules ready 42", Assert.Single(capture.Frames).Document.QuerySelector("#result")!.TextContent);
        Assert.DoesNotContain(capture.Resources, resource => resource.Url.AbsolutePath == "/wrong.js");
    }

    [Fact]
    public async Task DetachingAChildFrameCancelsItsPendingModuleEvaluation() {
        Uri frame = new(Page, "/frames/pending.html");
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = """
                <script>
                    addEventListener('message', event => {
                        if (event.data === 'module-started') {
                            document.querySelector('iframe').remove();
                            document.body.dataset.frameRemoved = 'yes';
                        }
                    });
                </script>
                <iframe src='/frames/pending.html'></iframe>
                """,
            Resources = new[] {
                HtmlRuntimeResource.FromText(frame, """
                    <script type='module'>
                        parent.postMessage('module-started', '*');
                        await new Promise(() => {});
                    </script>
                    """, "text/html; charset=utf-8")
            }
        });

        await session.WaitForAsync("document.body.dataset.frameRemoved === 'yes'");

        Assert.Empty((await session.CaptureAsync()).Frames);
        Assert.Equal(42, (await session.EvaluateAsync("40 + 2")).GetInt32());
    }

    [Fact]
    public async Task DetachingAChildFrameReleasesItsPendingModuleAcquisition() {
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var firstRequest = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var secondRequest = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        int slowRequests = 0;
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/frames/slow.js") {
                int request = Interlocked.Increment(ref slowRequests);
                if (request == 1) firstRequest.TrySetResult();
                if (request == 2) secondRequest.TrySetResult();
                await release.Task.WaitAsync(TimeSpan.FromSeconds(10), token);
                return RuntimeHttpFixture.Reply.Text("export const value=42");
            }
            return RuntimeHttpFixture.Reply.Text("not found", "text/plain", status: 404);
        });
        IHtmlRuntimeSession? session = null;
        try {
            Task<IHtmlRuntimeSession> opening = Runtime().OpenTrustedAsync(new() {
                DocumentUrl = server.Origin,
                ResourcePolicy = new() { AllowNetwork = true },
                Resources = new[] {
                    HtmlRuntimeResource.FromText(new Uri(server.Origin, "/frames/pending.html"),
                        "<script type='module' src='./pending.js'></script>", "text/html; charset=utf-8"),
                    HtmlRuntimeResource.FromText(new Uri(server.Origin, "/frames/pending.js"),
                        "const pending=import('./slow.js');setTimeout(()=>parent.postMessage('detach-pending-module','*'),250);await pending",
                        "text/javascript")
                },
                Html = """
                    <script>
                        addEventListener('message', event => {
                            if (event.data === 'detach-pending-module') {
                                document.querySelector('iframe').remove();
                                document.body.dataset.frameRemoved = 'yes';
                            }
                        });
                    </script>
                    <iframe src='/frames/pending.html'></iframe>
                    """
            });

            Task firstOrOpen = await Task.WhenAny(firstRequest.Task, opening, Task.Delay(TimeSpan.FromSeconds(5)));
            if (firstOrOpen == opening) session = await opening;
            Assert.True(firstOrOpen == firstRequest.Task,
                "The pending module request did not start. Requests: " + string.Join(", ", server.Requests));
            session = await opening.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.Equal("yes", (await session.EvaluateAsync("document.body.dataset.frameRemoved")).GetString());

            await session.ExecuteAsync("window.rootLoaded=false;import('/frames/slow.js').then(module=>rootLoaded=module.value===42)");
            await secondRequest.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.Equal(2, Volatile.Read(ref slowRequests));

            release.TrySetResult();
            await session.WaitForAsync("rootLoaded===true");
        } finally {
            release.TrySetResult();
            if (session != null) await session.DisposeAsync();
        }
    }

    [Fact]
    public async Task ChildAndRootRealmsShareTheModuleSourceCountBudget() {
        Uri frame = new(Page, "/frames/budget.html");
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            MaxModuleCount = 1,
            Html = "<script type='module'>export const root=1</script><iframe src='/frames/budget.html'></iframe>",
            Resources = new[] {
                HtmlRuntimeResource.FromText(frame,
                    "<script type='module'>export const child=1</script>", "text/html; charset=utf-8")
            }
        }));

        Assert.Contains("module source count budget", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task WrongDefaultModuleTypeDoesNotPoisonALaterJsonImportOfTheSameUrl() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Resources = new[] { Source("/settings.json", "{\"value\":42}", "text/json") },
            Scripts = new[] {
                "window.failure='';window.result=null;import('/settings.json').catch(error=>failure=String(error)).then(()=>import('/settings.json',{with:{type:'json'}})).then(module=>result=module.default.value)"
            }
        });

        await session.WaitForAsync("result!==null");

        Assert.Contains("JavaScript MIME", (await session.EvaluateAsync("failure")).GetString());
        Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
        Assert.Single((await session.CaptureAsync()).Resources, resource => resource.Url.AbsolutePath == "/settings.json");
    }

    [Theory]
    [InlineData("application/json", "{\"value\":42}", true)]
    [InlineData("text/json; charset=utf-8", "{\"value\":42}", true)]
    [InlineData("application/problem+json", "{\"value\":42}", true)]
    [InlineData("text/javascript", "{\"value\":42}", false)]
    [InlineData("not-a-mime+json", "{\"value\":42}", false)]
    [InlineData("application/json", "{invalid", false)]
    public async Task JsonModulesRequireJsonMimeAndValidJson(string contentType, string source, bool succeeds) {
        var request = new HtmlScriptRequest {
            DocumentUrl = Page,
            Html = "<script type='module'>import value from '/data.json' with {type:'json'};window.result=value.value</script>",
            Resources = new[] { Source("/data.json", source, contentType) }
        };

        if (succeeds) {
            await using IHtmlRuntimeSession session = await Runtime().OpenTrustedAsync(request);
            Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
        } else {
            await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().OpenTrustedAsync(request));
        }
    }

    [Fact]
    public async Task DynamicMalformedJsonModuleRejectsWithSyntaxError() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Resources = new[] { Source("/invalid.json", "{invalid", "application/json") },
            Scripts = new[] { "window.jsonError='';import('/invalid.json',{with:{type:'json'}}).catch(error=>jsonError=error.name)" }
        });

        await session.WaitForAsync("jsonError!==''");

        Assert.Equal("SyntaxError", (await session.EvaluateAsync("jsonError")).GetString());
    }

    [Fact]
    public async Task MultipleImportMapsAddRulesWithoutReplacingEarlierMappings() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = """
                <script type='importmap'>{"imports":{"dep":"/first.js"}}</script>
                <script type='importmap'>{"imports":{"dep":"/replacement.js","extra":"/extra.js"}}</script>
                <script type='module'>
                    import {value} from 'dep';
                    import {other} from 'extra';
                    window.result=value+other;
                    window.resolved=[import.meta.resolve('dep'),import.meta.resolve('./relative.js')];
                </script>
                """,
            Resources = new[] {
                Source("/first.js", "export const value=40"),
                Source("/replacement.js", "export const value=1"),
                Source("/extra.js", "export const other=2")
            }
        });

        Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
        Assert.Equal("https://modules.example/first.js,https://modules.example/nested/relative.js",
            (await session.EvaluateAsync("resolved.join(',')")).GetString());
        Assert.DoesNotContain((await session.CaptureAsync()).Resources, resource => resource.Url.AbsolutePath == "/replacement.js");
    }

    [Fact]
    public async Task LaterImportMapsPreserveResolvedSpecifiersAndCanAddFreshOnes() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = "<script type='importmap'>{\"imports\":{\"dep\":\"/first.js\"}}</script><script type='module'>import {value} from 'dep';window.before=value</script>",
            Resources = new[] { Source("/first.js", "export const value=40"), Source("/replacement.js", "export const value=1"), Source("/extra.js", "export const other=2") }
        });

        await session.ExecuteAsync("""
            const map=document.createElement('script');
            map.type='importmap';
            map.textContent=JSON.stringify({imports:{dep:'/replacement.js',extra:'/extra.js'}});
            document.head.append(map);
            window.after=null;
            Promise.all([import('dep'),import('extra')]).then(([dep,extra])=>after=dep.value+extra.other);
            """);
        await session.WaitForAsync("after!==null");

        Assert.True((await session.EvaluateAsync("before===40 && after===42")).GetBoolean());
    }

    [Fact]
    public async Task LaterImportMapPrefixesCannotChangeResolvedUrlsButCanAddMoreSpecificRules() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Resources = new[] {
                Source("/app/helper.js", "export const value=40"),
                Source("/replacement/helper.js", "export const value=1"),
                Source("/fresh.js", "export const value=2")
            },
            Scripts = new[] { "window.before=null;import('/app/helper.js').then(module=>before=module.value)" }
        });
        await session.WaitForAsync("before!==null");

        await session.ExecuteAsync("""
            const map=document.createElement('script');
            map.type='importmap';
            map.textContent=JSON.stringify({imports:{'/app/':'/replacement/','/app/helper.js/more':'/fresh.js'}});
            document.head.append(map);
            window.after=null;
            Promise.all([import('/app/helper.js'),import('/app/helper.js/more')]).then(([oldModule,fresh])=>after=oldModule.value+fresh.value);
            """);
        await session.WaitForAsync("after!==null");

        Assert.True((await session.EvaluateAsync("before===40 && after===42")).GetBoolean());
        Assert.DoesNotContain((await session.CaptureAsync()).Resources, resource => resource.Url.AbsolutePath == "/replacement/helper.js");
    }

    [Fact]
    public async Task ModulePreloadFetchIsReusedAndDoesNotDelayDocumentLoad() {
        const string source = "export const value=42";
        int requests = 0;
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/preloaded.js") {
                Interlocked.Increment(ref requests);
                started.TrySetResult();
                await release.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text(source, "text/javascript");
            }
            return RuntimeHttpFixture.Reply.Text("ok");
        });
        var opening = Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = $"<script>window.preloadLoaded=false</script><link rel='modulepreload' href='/preloaded.js' integrity='{Integrity(source)}' onload='preloadLoaded=true'><script>window.loaded=document.readyState</script>"
        });
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.Equal(1, Volatile.Read(ref requests));
        await using var session = await opening.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.False((await session.EvaluateAsync("preloadLoaded")).GetBoolean());
        await session.ExecuteAsync("import('/preloaded.js').then(module=>window.result=module.value)");
        release.SetResult();
        await session.WaitForAsync("window.result===42 && preloadLoaded");

        Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
        Assert.Equal(1, Volatile.Read(ref requests));
    }

    [Fact]
    public async Task ModulePreloadIntegrityFailurePoisonsTheModuleMapEntry() {
        const string source = "window.executed=true;export const value=42";
        string wrongIntegrity = Integrity("different source");
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = $"<link rel='modulepreload' href='/module.js' integrity='{wrongIntegrity}'><script type='module'>import '/module.js'</script>",
            Resources = new[] { Source("/module.js", source) }
        }));

        Assert.Contains("integrity", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ModulePreloadUsesImportMapIntegrityWhenTheLinkHasNoIntegrityAttribute() {
        const string source = "window.executed=true;export const value=42";
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = $"<script type='importmap'>{{\"integrity\":{{\"/module.js\":\"{Integrity("different source")}\"}}}}</script><link rel='modulepreload' href='/module.js'><script type='module'>import '/module.js'</script>",
            Resources = new[] { Source("/module.js", source) }
        }));

        Assert.Contains("integrity", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ModulePreloadAndExternalRootShareOneTransportRequest() {
        int requests = 0;
        await using var server = new RuntimeHttpFixture((path, _) => {
            if (path == "/main.js") Interlocked.Increment(ref requests);
            return Task.FromResult(RuntimeHttpFixture.Reply.Text("window.result=42;export const value=42", "text/javascript"));
        });
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = "<link rel='modulepreload' href='/main.js'><script type='module' src='/main.js'></script>"
        });

        Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
        Assert.Equal(1, Volatile.Read(ref requests));
    }

    [Fact]
    public async Task ExternalModuleScriptIntegrityIsEnforcedBeforeEvaluation() {
        const string source = "window.executed=true;export const value=42";
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = $"<script type='module' src='/module.js' integrity='{Integrity("different source")}'></script>",
            Resources = new[] { Source("/module.js", source) }
        }));

        Assert.Contains("integrity", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ExternalModuleScriptWithValidIntegrityRemainsReadableAndEvaluates() {
        const string source = "\uFEFFwindow.result=42;export const value=42";
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = $"<script type='module' src='/module.js' integrity='{Integrity(source)}'></script>",
            Resources = new[] { Source("/module.js", source) }
        });

        Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
    }

    [Fact]
    public async Task ExternalModuleUsesImportMapIntegrityWhenTheElementAttributeIsAbsent() {
        const string source = "window.executed=true;export const value=42";
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = $"<script type='importmap'>{{\"integrity\":{{\"/module.js\":\"{Integrity("different source")}\"}}}}</script><script type='module' src='/module.js'></script>",
            Resources = new[] { Source("/module.js", source) }
        }));

        Assert.Contains("integrity", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task EmptyElementIntegrityOverridesImportMapIntegrityForAnExternalModule() {
        const string source = "window.result=42;export const value=42";
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = $"<script type='importmap'>{{\"integrity\":{{\"/module.js\":\"{Integrity("different source")}\"}}}}</script><script type='module' src='/module.js' integrity=''></script>",
            Resources = new[] { Source("/module.js", source) }
        });

        Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
    }

    [Fact]
    public async Task LaterImportMapDoesNotChangeIntegritySelectedForAnInflightExternalModule() {
        const string source = "window.result=42;export const value=42";
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/module.js") {
                await release.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text(source, "text/javascript");
            }
            if (path == "/release") release.TrySetResult();
            return RuntimeHttpFixture.Reply.Text("ok", "text/plain");
        });
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = $"<script type='module' src='/module.js'></script><script type='importmap'>{{\"integrity\":{{\"/module.js\":\"{Integrity("different source")}\"}}}}</script><script>fetch('/release')</script>"
        });

        Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
    }

    [Fact]
    public async Task ExternalModuleIntegrityIsSnapshottedBeforeItsResponseArrives() {
        const string source = "window.result=42;export const value=42";
        var mutated = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/module.js") {
                await mutated.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text(source, "text/javascript");
            }
            if (path == "/mutated") mutated.TrySetResult();
            return RuntimeHttpFixture.Reply.Text("ok", "text/plain");
        });
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = $"<script id='module' type='module' src='/module.js' integrity='{Integrity(source)}'></script><script>document.querySelector('#module').integrity='{Integrity("wrong")}';fetch('/mutated')</script>"
        });

        Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
    }

    [Fact]
    public async Task MalformedStrongerIntegrityDigestDoesNotFallBackToWeakerMetadata() {
        const string source = "window.executed=true;export const value=42";
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = $"<script type='module' src='/module.js' integrity='sha512-!!! {Integrity(source)}'></script>",
            Resources = new[] { Source("/module.js", source) }
        }));

        Assert.Contains("integrity", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DistinctIntegrityMetadataForOneModuleIsBounded() {
        const string source = "window.loads=(window.loads||0)+1;export const value=42";
        string integrity = Integrity(source);
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => Runtime().OpenTrustedAsync(new() {
            MaxModuleIntegrityMetadataCharacters = 0
        }));
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            MaxModuleIntegrityMetadataCharacters = integrity.Length + 2,
            Html = $"<script type='module' src='/module.js' integrity='{integrity}?a'></script><script type='module' src='/module.js' integrity='{integrity}?b'></script>",
            Resources = new[] { Source("/module.js", source) }
        }));

        Assert.Contains("metadata", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("budget", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ImportMapIntegrityAppliesToImportedModuleSources() {
        const string source = "export const value=42";
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = $"<script type='importmap'>{{\"integrity\":{{\"/verified.js\":\"{Integrity(source)}\"}}}}</script><script type='module'>import {{value}} from '/verified.js';window.result=value</script>",
            Resources = new[] { Source("/verified.js", source) }
        });

        Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
    }

    [Fact]
    public async Task ReplacingModulePreloadHrefCancelsAndEvictsThePendingDownload() {
        var oldStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var oldRequestedAgain = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseFirst = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var newStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        int oldRequests = 0;
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/old.js") {
                oldStarted.TrySetResult();
                if (Interlocked.Increment(ref oldRequests) == 1) await releaseFirst.Task.WaitAsync(token);
                else oldRequestedAgain.TrySetResult();
            }
            if (path == "/replace.js") {
                await oldStarted.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("document.querySelector('#preload').href='/new.js'", "text/javascript");
            }
            if (path == "/new.js") newStarted.TrySetResult();
            return RuntimeHttpFixture.Reply.Text("export const value=42", "text/javascript");
        });
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = "<link id='preload' rel='modulepreload' href='/old.js'><script src='/replace.js'></script>"
        });

        await oldStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        await newStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
        await session.ExecuteAsync("import('/old.js').then(module=>window.reloaded=module.value)");
        await oldRequestedAgain.Task.WaitAsync(TimeSpan.FromSeconds(5));
        releaseFirst.TrySetResult();
        await session.WaitForAsync("window.reloaded===42");
        Assert.Equal(2, Volatile.Read(ref oldRequests));
    }

    [Fact]
    public async Task ReplacingModulePreloadHrefDoesNotCancelAnAttachedModuleImport() {
        const string source = "export const value=42";
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        int requests = 0;
        await using var server = new RuntimeHttpFixture(async (path, token) => {
            if (path == "/shared.js") {
                Interlocked.Increment(ref requests);
                started.TrySetResult();
                await release.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text(source, "text/javascript");
            }
            return RuntimeHttpFixture.Reply.Text("export const replacement=true", "text/javascript");
        });
        var opening = Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin,
            ResourcePolicy = new() { AllowNetwork = true },
            Html = "<link id='preload' rel='modulepreload' href='/shared.js'>"
        });
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        await using var session = await opening.WaitAsync(TimeSpan.FromSeconds(5));

        await session.ExecuteAsync("window.result=null;window.failure=null;import('/shared.js').then(module=>result=module.value).catch(error=>failure=String(error))");
        await session.ExecuteAsync("document.querySelector('#preload').href='/replacement.js'");
        release.SetResult();
        await session.WaitForAsync("result!==null || failure!==null");

        Assert.Equal(42, (await session.EvaluateAsync("result")).GetInt32());
        Assert.Equal(1, Volatile.Read(ref requests));
    }

    [Fact]
    public async Task CyclicGraphsRetainLiveBindingsAndRepeatedImportsShareOneNamespace() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl=Page, Resources=new[] {
                Source("/a.js", "import {b} from './b.js';export const a=20;export function total(){return a+b};globalThis.loads=(globalThis.loads||0)+1"),
                Source("/b.js", "import {a} from './a.js';export const b=22;export function read(){return a}")
            }, Scripts=new[] { "window.result=null;Promise.all([import('/a.js'),import('/a.js')]).then(([a,b])=>result={same:a===b,total:a.total(),loads})" }
        });
        await session.WaitForAsync("result!==null");
        Assert.True((await session.EvaluateAsync("result.same && result.total===42 && result.loads===1")).GetBoolean());
    }

    [Theory]
    [InlineData("file:///officeimo-do-not-read.mjs")]
    [InlineData("data:text/javascript,export default 1")]
    [InlineData("https://not-authorized.example/a.js")]
    [InlineData("unmapped-package")]
    public async Task DynamicImportPolicyFailuresAreCatchable(string specifier) {
        await using var session = await Runtime().OpenTrustedAsync(new() { DocumentUrl=Page });
        await session.ExecuteAsync("window.rejected=false;import(" + System.Text.Json.JsonSerializer.Serialize(specifier) + ").catch(()=>rejected=true)");
        await session.WaitForAsync("rejected===true");
        Assert.NotNull(await session.CaptureAsync());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NonJavaScriptMimeTypesAreRejectedForRootAndImportedModules(bool root) {
        var request = new HtmlScriptRequest { DocumentUrl=Page, Resources=new[] {Source("/bad.js","window.executed=true;export const value=1","text/plain")} };
        if (root) {
            request.Html="<script type='module' src='/bad.js'></script>";
            var error=await Assert.ThrowsAsync<HtmlScriptRuntimeException>(()=>Runtime().OpenTrustedAsync(request));
            Assert.Contains("MIME",error.Message);
        } else {
            await using var session=await Runtime().OpenTrustedAsync(request);
            await session.ExecuteAsync("window.failure='';import('/bad.js').catch(e=>failure=String(e))");
            await session.WaitForAsync("failure!==''");
            Assert.Contains("MIME",(await session.EvaluateAsync("failure")).GetString());
            Assert.Equal("undefined",(await session.EvaluateAsync("typeof executed")).GetString());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CrossOriginModulesRequireCorsForBothRootsAndDescendants(bool root) {
        var remote = new Uri("https://cdn.example/module.js");
        var request = new HtmlScriptRequest {
            DocumentUrl=Page, ResourcePolicy=new() { AllowedOrigins=new[]{new Uri("https://cdn.example/")} },
            Resources=new[]{HtmlRuntimeResource.FromText(remote,"window.executed=true;export const value=42","text/javascript")}
        };
        if(root) {
            request.Html="<script type='module' src='"+remote.AbsoluteUri+"'></script>";
            await Assert.ThrowsAsync<HtmlScriptRuntimeException>(()=>Runtime().OpenTrustedAsync(request));
        } else {
            await using var session=await Runtime().OpenTrustedAsync(request);
            await session.ExecuteAsync("window.rejected=false;import('"+remote.AbsoluteUri+"').catch(()=>rejected=true)");
            await session.WaitForAsync("rejected");
            Assert.Equal("undefined",(await session.EvaluateAsync("typeof executed")).GetString());
        }
    }

    [Fact]
    public async Task AuthorizedCorsModulesCanLoadAndRetainTheirOwnUrl() {
        var remote=new Uri("https://cdn.example/module.js");
        await using var session=await Runtime().OpenTrustedAsync(new() {
            DocumentUrl=Page, ResourcePolicy=new() {AllowedOrigins=new[]{new Uri("https://cdn.example/")}},
            Html="<script type='module' src='"+remote.AbsoluteUri+"'></script>",
            Resources=new[]{new HtmlRuntimeResource(remote,Encoding.UTF8.GetBytes("window.result=import.meta.url"),"text/javascript",headers:new Dictionary<string,string>{{"Access-Control-Allow-Origin","*"}})}
        });
        Assert.Equal(remote.AbsoluteUri,(await session.EvaluateAsync("result")).GetString());
    }

    [Fact]
    public async Task ModuleCountAndImportAttributeFailuresRemainCatchable() {
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(()=>Runtime().OpenTrustedAsync(new(){MaxModuleCount=0}));
        await using var session=await Runtime().OpenTrustedAsync(new() {
            DocumentUrl=Page, MaxModuleCount=1,
            Resources=new[]{Source("/a.js","export const value=1"),Source("/b.js","export const value=2")}
        });
        await session.ExecuteAsync("window.failure='';import('/a.js').then(()=>import('/b.js')).catch(e=>failure=String(e))");
        await session.WaitForAsync("failure!==''");
        Assert.Contains("count budget",(await session.EvaluateAsync("failure")).GetString());
        await session.ExecuteAsync("window.attributeError='';import('/a.js',{with:{type:'json'}}).catch(e=>attributeError=e.name)");
        await session.WaitForAsync("attributeError!==''");
        Assert.Equal("TypeError",(await session.EvaluateAsync("attributeError")).GetString());
        await session.ExecuteAsync("window.unknownAttribute='';import('/a.js',{with:{flavor:'json'}}).catch(e=>unknownAttribute=e.name)");
        await session.WaitForAsync("unknownAttribute!==''");
        Assert.Equal("TypeError", (await session.EvaluateAsync("unknownAttribute")).GetString());
        await session.ExecuteAsync("window.unsupportedType='';import('/a.js',{with:{type:'text'}}).catch(e=>unsupportedType=e.name)");
        await session.WaitForAsync("unsupportedType!==''");
        Assert.Equal("TypeError", (await session.EvaluateAsync("unsupportedType")).GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExternalClassicScriptsResolveDynamicImportsFromTheirOwnUrl(bool redirected) {
        await using var session=await Runtime().OpenTrustedAsync(new() {
            DocumentUrl=Page, Html="<script src='" + (redirected ? "/entry.js" : "/app/classic.js") + "'></script>",
            Resources=new[]{
                new HtmlRuntimeResource(new Uri(Page, "/entry.js"), Array.Empty<byte>(), "text/javascript", 302,
                    headers: new Dictionary<string,string>{{"Location","/app/classic.js"}}),
                Source("/app/classic.js","import('./dep.js').then(m=>window.result=m.value)"),Source("/app/dep.js","export const value=42")}
        });
        await session.WaitForAsync("window.result===42");
    }

    [Fact]
    public async Task NetworkModuleCompletionAdvancesWithoutHostPolling() {
        var release=new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var observed=new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server=new RuntimeHttpFixture((_,_)=>Task.FromResult(RuntimeHttpFixture.Reply.Text("ok")));
        server.RespondToRequest=async (request,token)=>{
            if(request.Path=="/delayed.js") {
                await release.Task.WaitAsync(token);
                return RuntimeHttpFixture.Reply.Text("export const value=42");
            }
            observed.TrySetResult(Encoding.UTF8.GetString(request.Body));
            return RuntimeHttpFixture.Reply.Text("ok");
        };
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=server.Origin,ResourcePolicy=new(){AllowNetwork=true}});
        await session.ExecuteAsync("import('/delayed.js').then(m=>fetch('/observed',{method:'POST',body:String(m.value)}))");
        release.SetResult();
        Assert.Equal("42",await observed.Task.WaitAsync(TimeSpan.FromSeconds(10)));
    }

    [Fact]
    public async Task RedirectedModulesResolveChildrenFromTheFinalResponseUrl() {
        await using var session=await Runtime().OpenTrustedAsync(new() {
            DocumentUrl=Page,
            Resources=new[]{
                new HtmlRuntimeResource(new Uri(Page,"/entry.js"),Array.Empty<byte>(),"text/javascript",302,headers:new Dictionary<string,string>{{"Location","/moved/main.js"}}),
                Source("/moved/main.js","import {value} from './dep.js';window.result=value;window.moduleUrl=import.meta.url"),
                Source("/moved/dep.js","export const value=42")
            },Scripts=new[]{"import('/entry.js')"}
        });
        await session.WaitForAsync("window.result===42");
        Assert.Equal("https://modules.example/moved/main.js",(await session.EvaluateAsync("moduleUrl")).GetString());
    }

    [Fact]
    public async Task FailedModuleLoadsAreCachedAndCountTowardTheSourceBudget() {
        int requests = 0;
        await using var server = new RuntimeHttpFixture((_, _) => {
            Interlocked.Increment(ref requests);
            return Task.FromResult(RuntimeHttpFixture.Reply.Text("not JavaScript", "text/plain"));
        });
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin, MaxModuleCount = 1, ResourcePolicy = new() { AllowNetwork = true }
        });
        await session.ExecuteAsync("window.failures=[];import('/bad.js').catch(e=>failures.push(String(e))).then(()=>import('/bad.js')).catch(e=>failures.push(String(e))).then(()=>import('/other.js')).catch(e=>failures.push(String(e)))");
        await session.WaitForAsync("failures.length===3");
        Assert.Equal(1, Volatile.Read(ref requests));
        Assert.Contains("MIME", (await session.EvaluateAsync("failures[1]")).GetString());
        Assert.Contains("count budget", (await session.EvaluateAsync("failures[2]")).GetString());
    }

    [Fact]
    public async Task CancellationInterruptsPendingTopLevelAwaitAfterModuleExecutionStarts() {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture((_, _) => {
            started.TrySetResult();
            return Task.FromResult(RuntimeHttpFixture.Reply.Text("ok"));
        });
        using var cancellation = new CancellationTokenSource();
        var opening = Runtime().OpenTrustedAsync(new() {
            DocumentUrl = server.Origin, Timeout = TimeSpan.FromSeconds(30), ResourcePolicy = new() { AllowNetwork = true },
            Html = "<script type='module'>await fetch('/started');await new Promise(()=>{})</script>"
        }, cancellation.Token);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(10));
        cancellation.Cancel();
        var error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => opening.WaitAsync(TimeSpan.FromSeconds(5)));
        Assert.Equal(cancellation.Token, error.CancellationToken);
    }

    [Theory]
    [InlineData("<script type='module'>export const a=1</script><script type='module'>export const b=2</script>", "count budget")]
    [InlineData("<script type='module' src='/a.js' crossorigin='use-credentials'></script>", "Credentialed")]
    public async Task UnsupportedDocumentModuleBoundariesFailExplicitly(string html, string message) {
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page, Html = html, MaxModuleCount = 1, Resources = new[] { Source("/a.js", "export const value=42") }
        }));
        Assert.Contains(message, error.Message);
    }

    [Theory]
    [InlineData(false, "/final.js", "#original")]
    [InlineData(true, "/final.js#replacement", "#replacement")]
    [InlineData(false, "/final.js#", "#")]
    public async Task ModuleMetadataRetainsRedirectFragmentsWhileFetchResponseUrlOmitsThem(bool root, string location, string fragment) {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page,
            Html = root ? "<script type='module' src='/entry.js#original'></script>" : "",
            Resources = new[] {
                new HtmlRuntimeResource(new Uri(Page, "/entry.js"), Array.Empty<byte>(), "text/javascript", 302,
                    headers: new Dictionary<string, string> { { "Location", location } }),
                Source("/final.js", "window.moduleUrl=import.meta.url")
            },
            Scripts = root ? Array.Empty<string>() : new[] { "import('/entry.js#original')" }
        });
        await session.WaitForAsync("window.moduleUrl!==undefined");
        Assert.Equal("https://modules.example/final.js" + fragment, (await session.EvaluateAsync("moduleUrl")).GetString());
        await session.ExecuteAsync("fetch('/entry.js#original').then(r=>window.responseUrl=r.url)");
        await session.WaitForAsync("window.responseUrl!==undefined");
        Assert.Equal("https://modules.example/final.js", (await session.EvaluateAsync("responseUrl")).GetString());
    }

    [Fact]
    public async Task RefusingADedicatedWorkerPreservesTheSessionsModuleEngine() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = Page, Resources = new[] { Source("/a.js", "export const value=42") },
            Scripts = new[] { "try{new Worker('/worker.js')}catch(e){window.workerError=e.name};import('/a.js').then(m=>window.result=m.value)" }
        });
        await session.WaitForAsync("window.result===42");
        Assert.Equal("NotSupportedError", (await session.EvaluateAsync("workerError")).GetString());
    }
}
