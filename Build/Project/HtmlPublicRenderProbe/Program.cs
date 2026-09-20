using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Rendering;
using OfficeIMO.Pdf;

if (args.Length is < 4 or > 5 || args.Length == 5 && args[4] != "--live-acquisition")
    throw new ArgumentException("Usage: OfficeIMO.Html.PublicRenderProbe <full-sha256-image-id> <published-renderer-dll> <published-worker-dll> <new-output-directory> [--live-acquisition]");

string imageId = args[0];
string rendererPath = Path.GetFullPath(args[1]);
string workerPath = Path.GetFullPath(args[2]);
string outputDirectory = Path.GetFullPath(args[3]);
bool runLiveAcquisition = args.Length == 5;
if (Directory.Exists(outputDirectory) || File.Exists(outputDirectory))
    throw new IOException("The probe output directory must be new.");
Directory.CreateDirectory(outputDirectory);

string rendererSha256 = Digest(await File.ReadAllBytesAsync(rendererPath));
string workerSha256 = Digest(await File.ReadAllBytesAsync(workerPath));
string rendererFilesSha256 = HtmlPublicArtifactDigest.DirectorySha256(Path.GetDirectoryName(rendererPath)!);
string workerFilesSha256 = HtmlPublicArtifactDigest.DirectorySha256(Path.GetDirectoryName(workerPath)!);
HtmlPublicFontPackageIdentity fontPackage = HtmlPublicFontPackage.Load(Path.GetDirectoryName(rendererPath)!);
const string fixtureOrigin = "https://fixture.officeimo.invalid";
var cases = new List<ProbeCase> {
    new ProbeCase("malformed-markup", """
        <!doctype html><style>body{font:16px sans-serif}p{color:#0055aa}</style>
        <main><table><tr><td><p id=result>Before script
        <script>document.querySelector('#result').textContent='Recovered by script';</script>
        """, "document.querySelector('#result')?.textContent === 'Recovered by script'", 8 * 1024 * 1024,
        ExpectedVisibleText: "Recovered by script", ExpectBlueInk: true),
    new ProbeCase("responsive-picture", """
        <!doctype html><style>body{font:16px sans-serif}</style>
        <p>Responsive source ready</p>
        <picture>
          <source media="(max-width: 900px)" type="image/svg+xml" sizes="400px"
                  srcset="/responsive-1x.svg 400w, /responsive-2x.svg 800w, /responsive-3x.svg 1200w">
          <source media="(min-width: 901px)" type="image/svg+xml" srcset="/responsive-narrow.svg">
          <img src="/responsive-fallback.svg" width="180" height="80" alt="responsive fixture">
        </picture>
        """, "document.readyState === 'complete'", 8 * 1024 * 1024,
        ExpectedVisibleText: "Responsive source ready", ExpectBlueInk: true,
        Resources: new Dictionary<string, ProbeResource>(StringComparer.Ordinal) {
            [$"{fixtureOrigin}/responsive-2x.svg"] = new("""
                <svg xmlns="http://www.w3.org/2000/svg" width="180" height="80" viewBox="0 0 180 80">
                  <rect width="180" height="80" fill="#0055aa"/>
                </svg>
                """, "image/svg+xml")
        }, ExpectedDiscoveryRounds: [[ $"{fixtureOrigin}/responsive-2x.svg" ]], DevicePixelRatio: 2D),
    new ProbeCase("module-graph", """
        <!doctype html><style>body{font:16px sans-serif}#result{color:#0055aa}</style>
        <p id="result">Loading module graph</p><script type="module" src="/app/main.js"></script>
        """, "document.querySelector('#result')?.textContent === 'Module graph ready'", 8 * 1024 * 1024,
        ExpectedVisibleText: "Module graph ready", ExpectBlueInk: true,
        Resources: new Dictionary<string, ProbeResource>(StringComparer.Ordinal) {
            [$"{fixtureOrigin}/app/main.js"] = new("import { message } from './dep.js'; document.querySelector('#result').textContent = message;", "text/javascript"),
            [$"{fixtureOrigin}/app/dep.js"] = new("export const message = 'Module graph ready';", "text/javascript")
        }, ExpectedDiscoveryRounds: [
            [ $"{fixtureOrigin}/app/main.js" ],
            [ $"{fixtureOrigin}/app/dep.js" ]
        ]),
    new ProbeCase("import-map-graph", """
        <!doctype html><style>body{font:16px sans-serif}#result{color:#0055aa}</style>
        <script type="importmap">
          {"imports":{"pkg/":"/vendor/pkg/","theme":"/vendor/global-theme.js"},
           "scopes":{"/app/features/":{"theme":"/vendor/scoped-theme.js"}}}
        </script>
        <p id="result">Loading mapped module graph</p>
        <script type="module" src="/app/features/main.js"></script>
        """, "document.querySelector('#result')?.textContent === 'Scoped import map ready 42'", 8 * 1024 * 1024,
        ExpectedVisibleText: "Scoped import map ready 42", ExpectBlueInk: true,
        Resources: new Dictionary<string, ProbeResource>(StringComparer.Ordinal) {
            [$"{fixtureOrigin}/app/features/main.js"] = new("""
                import { add } from 'pkg/math.js';
                import { label } from 'theme';
                const details = await import('./details.js');
                document.querySelector('#result').textContent = `${label} ${add(20, details.delta)}`;
                """, "text/javascript"),
            [$"{fixtureOrigin}/vendor/pkg/math.js"] = new("export const add = (left, right) => left + right;", "text/javascript"),
            [$"{fixtureOrigin}/vendor/scoped-theme.js"] = new("export const label = 'Scoped import map ready';", "text/javascript"),
            [$"{fixtureOrigin}/app/features/details.js"] = new("export const delta = 22;", "text/javascript")
        }, ExpectedDiscoveryRounds: [
            [ $"{fixtureOrigin}/app/features/main.js" ],
            [ $"{fixtureOrigin}/vendor/pkg/math.js" ],
            [ $"{fixtureOrigin}/vendor/scoped-theme.js" ],
            [ $"{fixtureOrigin}/app/features/details.js" ]
        ]),
    new ProbeCase("dynamic-fetch-get", """
        <!doctype html><style>body{font:16px sans-serif}#result{color:#0055aa}</style>
        <p id="result">Loading dynamic data</p>
        <script>
          fetch('/api/report.json?view=summary#client')
            .then(response => response.json())
            .then(data => document.querySelector('#result').textContent = `Dynamic fetch ready ${data.total}`);
        </script>
        """, "document.querySelector('#result')?.textContent === 'Dynamic fetch ready 42'", 8 * 1024 * 1024,
        ExpectedVisibleText: "Dynamic fetch ready 42", ExpectBlueInk: true,
        Resources: new Dictionary<string, ProbeResource>(StringComparer.Ordinal) {
            [$"{fixtureOrigin}/api/report.json?view=summary"] = new("{\"total\":42}", "application/json")
        }, ExpectedDiscoveryRounds: [[ $"{fixtureOrigin}/api/report.json?view=summary" ]]),
    new ProbeCase("dynamic-xhr-get", """
        <!doctype html><style>body{font:16px sans-serif}#result{color:#0055aa}</style>
        <p id="result">Loading XHR data</p>
        <script>
          const request = new XMLHttpRequest();
          request.open('GET', '/api/xhr-report.json?view=summary#client');
          request.responseType = 'json';
          request.onload = () => document.querySelector('#result').textContent = `XHR ready ${request.response.total}`;
          request.onerror = () => document.querySelector('#result').textContent = 'XHR acquisition pending';
          request.onloadend = () => document.body.dataset.xhrSettled = 'yes';
          request.send();
        </script>
        """, "document.body.dataset.xhrSettled === 'yes'", 8 * 1024 * 1024,
        ExpectedVisibleText: "XHR ready 42", ExpectBlueInk: true,
        Resources: new Dictionary<string, ProbeResource>(StringComparer.Ordinal) {
            [$"{fixtureOrigin}/api/xhr-report.json?view=summary"] = new("{\"total\":42}", "application/json")
        }, ExpectedDiscoveryRounds: [[ $"{fixtureOrigin}/api/xhr-report.json?view=summary" ]]),
    new ProbeCase("dynamic-xhr-headered-get", """
        <!doctype html><style>body{font:16px sans-serif}#result{color:#0055aa}</style>
        <p id="result">Loading header-varying XHR</p>
        <script>
          const request = new XMLHttpRequest();
          request.open('GET', '/api/header-varying.json');
          request.setRequestHeader('X-Variant', 'private');
          request.onload = () => document.querySelector('#result').textContent = `Headered XHR ready ${request.responseText}`;
          request.onerror = () => document.querySelector('#result').textContent = 'Headered XHR failed';
          request.onloadend = () => document.body.dataset.xhrSettled = 'yes';
          request.send();
        </script>
        """, "document.body.dataset.xhrSettled === 'yes'", 8 * 1024 * 1024,
        ExpectedVisibleText: "Headered XHR ready 42", ExpectBlueInk: true,
        DynamicResponses: new Dictionary<string, ProbeDynamicResource>(StringComparer.Ordinal) {
            [$"GET {fixtureOrigin}/api/header-varying.json #1"] = new("42", "text/plain", ExpectedHeaders:
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["X-Variant"] = "private" })
        }, ExpectedDiscoveryRounds: [[ $"GET {fixtureOrigin}/api/header-varying.json #1" ]]),
    new ProbeCase("dynamic-post-occurrences", """
        <!doctype html><style>body{font:16px sans-serif}#result{color:#0055aa}</style>
        <p id="result">Posting dynamic data</p>
        <script>
          (async()=>{
            const options={method:'POST',headers:{'Content-Type':'application/json','X-Variant':'blue'},body:'{"value":42}'};
            const first=await (await fetch('/api/submit',options)).text();
            const second=await (await fetch('/api/submit',options)).text();
            document.querySelector('#result').textContent=`Dynamic POST ready ${first}/${second}`;
          })();
        </script>
        """, "document.querySelector('#result')?.textContent === 'Dynamic POST ready first/second'", 8 * 1024 * 1024,
        ExpectedVisibleText: "Dynamic POST ready first/second", ExpectBlueInk: true,
        DynamicResponses: new Dictionary<string, ProbeDynamicResource>(StringComparer.Ordinal) {
            [$"POST {fixtureOrigin}/api/submit #1"] = new("first", "text/plain", "{\"value\":42}",
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["Content-Type"] = "application/json", ["X-Variant"] = "blue" }),
            [$"POST {fixtureOrigin}/api/submit #2"] = new("second", "text/plain", "{\"value\":42}",
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["Content-Type"] = "application/json", ["X-Variant"] = "blue" })
        }, ExpectedDiscoveryRounds: [
            [ $"POST {fixtureOrigin}/api/submit #1" ],
            [ $"POST {fixtureOrigin}/api/submit #2" ]
        ]),
    new ProbeCase("dynamic-post-redirect", """
        <!doctype html><style>#result{color:#0055aa}</style><p id=result>Loading</p>
        <script>fetch('/api/submit',{method:'POST',body:'once'}).then(r=>r.text()).then(t=>document.querySelector('#result').textContent=t)</script>
        """, "document.querySelector('#result')?.textContent === 'Redirected fetch ready'", 8 * 1024 * 1024,
        ExpectedVisibleText: "Redirected fetch ready", ExpectBlueInk: true,
        DynamicResponses: new Dictionary<string, ProbeDynamicResource>(StringComparer.Ordinal) {
            [$"POST {fixtureOrigin}/api/submit #1"] = new("", "text/plain", "once", Hops: new[] {
                new HtmlRuntimeFetchHop(new HtmlRuntimeResource(new Uri($"{fixtureOrigin}/api/submit"), [], "text/plain", 302,
                    headers: new Dictionary<string, string> { ["Location"] = "/api/result" })),
                new HtmlRuntimeFetchHop(HtmlRuntimeResource.FromText(new Uri($"{fixtureOrigin}/api/result"),
                    "Redirected fetch ready", "text/plain"))
            })
        }, ExpectedDiscoveryRounds: [[ $"POST {fixtureOrigin}/api/submit #1" ]]),
    new ProbeCase("dynamic-cross-origin-preflight", """
        <!doctype html><style>#result{color:#0055aa}</style><p id=result>Loading</p>
        <script>fetch('https://api.fixture.officeimo.invalid/data',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}',credentials:'omit'}).then(r=>r.text()).then(t=>document.querySelector('#result').textContent=t)</script>
        """, "document.querySelector('#result')?.textContent === 'CORS fetch ready'", 8 * 1024 * 1024,
        ExpectedVisibleText: "CORS fetch ready", ExpectBlueInk: true,
        AllowedOrigins: [new Uri("https://api.fixture.officeimo.invalid/")],
        DynamicResponses: new Dictionary<string, ProbeDynamicResource>(StringComparer.Ordinal) {
            ["POST https://api.fixture.officeimo.invalid/data #1"] = new("", "text/plain", "{}",
                new Dictionary<string, string> { ["Content-Type"] = "application/json" }, Hops: new[] {
                    new HtmlRuntimeFetchHop(new HtmlRuntimeResource(new Uri("https://api.fixture.officeimo.invalid/data"),
                        Encoding.UTF8.GetBytes("CORS fetch ready"), "text/plain", headers: new Dictionary<string, string> {
                            ["Access-Control-Allow-Origin"] = fixtureOrigin
                        }), new HtmlRuntimeResource(new Uri("https://api.fixture.officeimo.invalid/data"), [],
                        "text/plain", 204, headers: new Dictionary<string, string> {
                            ["Access-Control-Allow-Origin"] = fixtureOrigin,
                            ["Access-Control-Allow-Methods"] = "POST",
                            ["Access-Control-Allow-Headers"] = "content-type"
                        }))
                })
        }, ExpectedDiscoveryRounds: [[ "POST https://api.fixture.officeimo.invalid/data #1" ]]),
    new ProbeCase("frame-document", """
        <!doctype html><style>body{font:16px sans-serif}#outer{color:#222}</style>
        <script>addEventListener('message',event=>{if(event.data?.kind==='frame-ready')document.body.dataset.childMessage=event.data.value})</script>
        <p id="outer">Outer frame host ready</p>
        <iframe src="/frame/detail.html"></iframe>
        """, "document.querySelector('iframe')?.contentDocument?.querySelector('#inside')?.textContent === 'Frame realm ready' && document.body.dataset.childEvent === 'ran' && document.body.dataset.childInline === 'ran' && document.body.dataset.childExternal === 'ran' && document.body.dataset.childMessage === 'ready'",
        8 * 1024 * 1024, ExpectedVisibleText: "Frame realm ready", ExpectMagentaArea: true,
        Resources: new Dictionary<string, ProbeResource>(StringComparer.Ordinal) {
            [$"{fixtureOrigin}/frame/detail.html"] = new("""
                <!doctype html><link rel="stylesheet" href="frame.css">
                <body onload="parent.document.body.dataset.childEvent='ran'">
                <p id="inside">Frame document ready</p>
                <script>document.querySelector('#inside').textContent='inline ran';parent.document.body.dataset.childInline='ran'</script>
                <script src="frame.js"></script>
                """, "text/html; charset=utf-8"),
            [$"{fixtureOrigin}/frame/frame.css"] = new("body{margin:0;background:#d1007f}#inside{color:white}", "text/css"),
            [$"{fixtureOrigin}/frame/frame.js"] = new("document.querySelector('#inside').textContent='Frame realm ready';parent.document.body.dataset.childExternal='ran';parent.postMessage({kind:'frame-ready',value:'ready'},'*')", "text/javascript")
        }, ExpectedDiscoveryRounds: [
            [ $"{fixtureOrigin}/frame/detail.html" ],
            [ $"{fixtureOrigin}/frame/frame.css", $"{fixtureOrigin}/frame/frame.js" ]
        ]),
    new ProbeCase("frame-module-json", """
        <!doctype html><style>body{font:16px sans-serif}#result{color:#0055aa}</style>
        <script>addEventListener('message',event=>{if(event.data?.kind==='frame-module')document.querySelector('#result').textContent=event.data.value})</script>
        <p id="result">Loading frame module graph</p>
        <iframe src="/modules/frame.html"></iframe>
        """, "document.querySelector('#result')?.textContent === 'Frame module graph ready 42'",
        8 * 1024 * 1024, ExpectedVisibleText: "Frame module graph ready 42", ExpectBlueInk: true,
        Resources: new Dictionary<string, ProbeResource>(StringComparer.Ordinal) {
            [$"{fixtureOrigin}/modules/frame.html"] = new("""
                <!doctype html><style>body{margin:0;background:#e8f2ff}#inside{color:#17324d}</style>
                <script type="importmap">{"imports":{"settings":"./settings.json"}}</script>
                <p id="inside">Loading child graph</p>
                <script type="module" src="./main.js"></script>
                """, "text/html; charset=utf-8"),
            [$"{fixtureOrigin}/modules/main.js"] = new("""
                import {label} from './dep.js';
                import settings from 'settings' with {type:'json'};
                const extra=await import('./extra.json',{with:{type:'json'}});
                const value=`${label} ${settings.value+extra.default.delta}`;
                document.querySelector('#inside').textContent=value;
                parent.postMessage({kind:'frame-module',value},'*');
                """, "text/javascript"),
            [$"{fixtureOrigin}/modules/dep.js"] = new("export const label='Frame module graph ready';", "text/javascript"),
            [$"{fixtureOrigin}/modules/settings.json"] = new("{\"value\":40}", "application/json"),
            [$"{fixtureOrigin}/modules/extra.json"] = new("{\"delta\":2}", "application/vnd.officeimo+json")
        }, ExpectedDiscoveryRounds: [
            [ $"{fixtureOrigin}/modules/frame.html" ],
            [ $"{fixtureOrigin}/modules/main.js" ],
            [ $"{fixtureOrigin}/modules/dep.js" ],
            [ $"{fixtureOrigin}/modules/settings.json" ],
            [ $"{fixtureOrigin}/modules/extra.json" ]
        ]),
    new ProbeCase("json-module-rejections", """
        <!doctype html><style>body{font:16px sans-serif}#result{color:#0055aa}</style>
        <p id="result">Checking JSON module failures</p>
        <img hidden src="/modules/invalid-mime.json" alt="">
        <img hidden src="/modules/invalid-json.json" alt="">
        <script type="module">
            const mimeError=await import('/modules/invalid-mime.json',{with:{type:'json'}}).catch(error=>error.name);
            const syntaxError=await import('/modules/invalid-json.json',{with:{type:'json'}}).catch(error=>error.name);
            document.querySelector('#result').textContent=`JSON rejected ${mimeError}/${syntaxError}`;
        </script>
        """, "document.querySelector('#result')?.textContent === 'JSON rejected TypeError/SyntaxError'",
        8 * 1024 * 1024, ExpectedVisibleText: "JSON rejected TypeError/SyntaxError", ExpectBlueInk: true,
        Resources: new Dictionary<string, ProbeResource>(StringComparer.Ordinal) {
            [$"{fixtureOrigin}/modules/invalid-mime.json"] = new("{\"value\":42}", "not-a-mime+json"),
            [$"{fixtureOrigin}/modules/invalid-json.json"] = new("{invalid", "application/json")
        }, ExpectedDiscoveryRounds: [[
            $"{fixtureOrigin}/modules/invalid-mime.json",
            $"{fixtureOrigin}/modules/invalid-json.json"
        ]]),
    new ProbeCase("resource-fanout", "<!doctype html>" + string.Concat(Enumerable.Range(0, 129)
        .Select(index => $"<script src='/asset-{index}.js'></script>")), "true", 8 * 1024 * 1024,
        ExpectedErrorKind: "HtmlScriptRuntimeException",
        ExpectedError: "The document exceeds the pilot's static resource discovery limit."),
    new ProbeCase("capture-output-budget", "<p>" + new string('X', 4096) + "</p>", "true", 1024,
        ExpectedErrorKind: "HtmlScriptRuntimeException", ExpectedError: "Captured data budget exceeded."),
    new ProbeCase("encoded-output-budget", "<p>Encoded output limit</p>", "true", 8 * 1024 * 1024,
        ExpectedErrorKind: "HtmlScriptRuntimeException",
        ExpectedError: "The isolated render output exceeds its byte budget.", MaxOutputBytesPerArtifact: 1),
    new ProbeCase("runaway-script", "<script>while(true){}</script>", "true", 8 * 1024 * 1024,
        ExpectedErrorKind: "TimeoutException", ExpectedError: "The runtime command exceeded its deadline."),
    new ProbeCase("runaway-frame-script", "<iframe src='/frame/runaway.html'></iframe>", "true", 8 * 1024 * 1024,
        ExpectedErrorKind: "TimeoutException", ExpectedError: "The runtime command exceeded its deadline.",
        Resources: new Dictionary<string, ProbeResource>(StringComparer.Ordinal) {
            [$"{fixtureOrigin}/frame/runaway.html"] = new("<script>while(true){}</script>", "text/html; charset=utf-8")
        }, ExpectedDiscoveryRounds: [[ $"{fixtureOrigin}/frame/runaway.html" ]])
};
ControlledAcquisitionCorpus acquisition = await ControlledAcquisitionCorpus.CreateAsync();
cases.AddRange(acquisition.RenderCases);
var results = new List<ProbeResult>();
foreach (ProbeCase fixture in cases) {
    ProbeResult result = await RunCaseAsync(fixture);
    results.Add(result);
    if (!result.ContainerRemoved) break;
}
IsolatedApplicationCorpus applications = results.All(result => result.ContainerRemoved)
    ? await IsolatedApplicationCorpus.RunAsync(imageId, rendererPath, workerPath, outputDirectory)
    : new IsolatedApplicationCorpus(Array.Empty<IsolatedApplicationProbeResult>());
LivePublicAcquisitionCorpus liveAcquisition = runLiveAcquisition
    ? await LivePublicAcquisitionCorpus.RunAsync()
    : new LivePublicAcquisitionCorpus(Array.Empty<LivePublicAcquisitionResult>());

await File.WriteAllTextAsync(Path.Combine(outputDirectory, "summary.json"), JsonSerializer.Serialize(new {
    imageId, rendererSha256, workerSha256, rendererFilesSha256, workerFilesSha256,
    fontPackage = new { fontPackage.Id, fontPackage.ManifestSha256, fontPackage.FilesSha256,
        fontCount = fontPackage.Fonts.Length, licenseCount = fontPackage.Licenses.Length },
    capturedAtUtc = DateTimeOffset.UtcNow, isolationPolicy =
        "rootless-podman;seccomp;cgroups-cpu-memory-pids;network-none;read-only;uid-65532;cap-drop-all;no-new-privileges;no-mounts",
    acquisitionCases = acquisition.Results,
    liveAcquisitionRequested = runLiveAcquisition,
    liveAcquisitionCases = liveAcquisition.Results,
    isolatedApplicationCases = applications.Results,
    cases = results
}, new JsonSerializerOptions { WriteIndented = true }));
if (acquisition.Results.Any(result => !result.Passed) || results.Any(result => !result.Passed)
    || applications.Results.Count != 3 || applications.Results.Any(result => !result.Passed)
    || runLiveAcquisition && liveAcquisition.Results.Any(result => !result.Passed))
    throw new InvalidOperationException("The isolated renderer probe failed; inspect summary.json.");
Console.WriteLine("Isolated renderer probe passed: " + results.Count + " render cases and " +
    acquisition.Results.Count + " acquisition cases, " + applications.Results.Count +
    " isolated application cases" + (runLiveAcquisition ? ", and 2 live acquisition cases." : "."));

async Task<ProbeResult> RunCaseAsync(ProbeCase fixture) {
    using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
    HtmlOciWorkerLease? lease = null;
    string? containerName = null;
    bool removed = false;
    bool unexpectedException = false;
    bool workerReportedError = false;
    string? errorKind = null, errorMessage = null;
    string? cleanupError = null;
    HtmlPublicRenderResponse? final = null;
    var discoveryRounds = new List<string[]>();
    var stopwatch = Stopwatch.StartNew();
    try {
        lease = await HtmlOciWorkerLease.StartAsync("podman", imageId, deadline.Token);
        containerName = lease.ContainerName;
        Process process = lease.Process;
        Task<string> stderr = DrainErrorAsync(process.StandardError.BaseStream, deadline.Token);
        var page = new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = fixture.Html,
            DocumentUrl = fixture.DocumentUrl ?? new Uri(fixtureOrigin + "/"),
            ReadyExpression = fixture.ReadyExpression,
            ResourcePolicy = new HtmlRuntimeResourcePolicy { AllowNetwork = false, MaxRequests = 64,
                AllowedOrigins = fixture.AllowedOrigins ?? Array.Empty<Uri>() },
            Timeout = TimeSpan.FromSeconds(10), SessionTimeout = TimeSpan.FromSeconds(30),
            MaxOutputCharacters = fixture.MaxOutputCharacters,
            DevicePixelRatio = fixture.DevicePixelRatio,
            ViewportWidth = 816D,
            ViewportHeight = 720D
        };
        page.FailOnFetchReplayDiscovery = true;
        await HtmlRuntimeProtocol.WriteAsync(process.StandardInput.BaseStream,
            new HtmlPublicRenderRequest { Page = page,
                MaxOutputBytesPerArtifact = fixture.MaxOutputBytesPerArtifact ?? 8L * 1024 * 1024 },
            24 * 1024 * 1024, deadline.Token);
        for (int round = 0; round <= 16; round++) {
            var response = await HtmlRuntimeProtocol.ReadAsync<HtmlPublicRenderResponse>(
                process.StandardOutput.BaseStream, 24 * 1024 * 1024, deadline.Token)
                ?? throw new IOException("The isolated renderer exited without a response.");
            VerifyIdentity(response);
            if (response.Error != null) {
                workerReportedError = true;
                errorKind = response.ErrorKind;
                errorMessage = response.Error;
                break;
            }
            if (response.DiscoveryComplete) {
                final = await HtmlRuntimeProtocol.ReadAsync<HtmlPublicRenderResponse>(
                    process.StandardOutput.BaseStream, 24 * 1024 * 1024, deadline.Token)
                    ?? throw new IOException("The isolated renderer did not send its final output.");
                VerifyIdentity(final);
                workerReportedError = final.Error != null;
                errorKind = final.ErrorKind;
                errorMessage = final.Error;
                break;
            }
            var batch = new List<HtmlRuntimeResource>();
            var fetchBatch = new List<HtmlRuntimeFetchReplay>();
            var discoveryRound = new List<string>();
            foreach (string candidate in response.DiscoveryUrls) {
                if (!Uri.TryCreate(candidate, UriKind.Absolute, out Uri? resourceUrl))
                    throw new IOException("The isolated renderer requested an invalid discovery URL: " + candidate);
                if (!string.IsNullOrEmpty(resourceUrl.Fragment))
                    throw new IOException("The isolated renderer requested a resource URL with a fragment: " + candidate);
                string canonicalUrl = resourceUrl.AbsoluteUri;
                discoveryRound.Add(canonicalUrl);
                if (fixture.Resources == null || !fixture.Resources.TryGetValue(canonicalUrl, out ProbeResource? resource))
                    throw new IOException("The isolated renderer requested an unexpected resource: " + candidate);
                batch.Add(HtmlRuntimeResource.FromText(resourceUrl, resource.Content, resource.ContentType));
            }
            foreach (HtmlRuntimeFetchDiscovery discovery in response.DiscoveryRequests) {
                HtmlRuntimeFetchRequest request = discovery.Request;
                string canonicalUrl = request.Url.AbsoluteUri;
                string identity = request.Method + " " + canonicalUrl + " #" + discovery.Occurrence;
                ProbeDynamicResource? expected = null;
                if (request.Method == "GET" && request.Headers.Count == 0 && request.BodyLength == 0 &&
                    fixture.Resources?.TryGetValue(canonicalUrl, out ProbeResource? simple) == true) {
                    expected = new ProbeDynamicResource(simple.Content, simple.ContentType);
                    discoveryRound.Add(canonicalUrl);
                } else {
                    discoveryRound.Add(identity);
                    if (fixture.DynamicResponses == null || !fixture.DynamicResponses.TryGetValue(identity, out expected))
                        throw new IOException("The isolated renderer requested an unexpected dynamic resource: " + identity);
                }
                if (expected.ExpectedBody != null && Encoding.UTF8.GetString(request.Body ?? Array.Empty<byte>()) != expected.ExpectedBody)
                    throw new IOException("The isolated renderer requested the wrong dynamic body: " + identity);
                foreach (var header in expected.ExpectedHeaders ?? new Dictionary<string, string>())
                    if (!request.Headers.TryGetValue(header.Key, out string? actual) || actual != header.Value)
                        throw new IOException("The isolated renderer requested the wrong dynamic header: " + identity + " " + header.Key);
                if (expected.Hops != null)
                    fetchBatch.Add(new HtmlRuntimeFetchReplay(request, discovery.Occurrence, expected.Hops));
                else {
                    var resource = HtmlRuntimeResource.FromText(request.Url, expected.Content, expected.ContentType);
                    fetchBatch.Add(new HtmlRuntimeFetchReplay(request, discovery.Occurrence, resource));
                }
            }
            discoveryRounds.Add(discoveryRound.ToArray());
            await HtmlRuntimeProtocol.WriteAsync(process.StandardInput.BaseStream,
                new HtmlPublicResourceBatch { Resources = batch.ToArray(), FetchReplays = fetchBatch.ToArray() },
                24 * 1024 * 1024, deadline.Token);
        }
        process.StandardInput.Close();
        await process.WaitForExitAsync(deadline.Token);
        string stderrText = await stderr;
        if (process.ExitCode != 0) throw new IOException("The isolated renderer exited " + process.ExitCode + ": " + stderrText);
        if (!workerReportedError && fixture.ExpectedDiscoveryRounds != null &&
            !DiscoveryRoundsEqual(discoveryRounds, fixture.ExpectedDiscoveryRounds))
            throw new IOException("The isolated renderer requested the wrong ordered discovery rounds: " +
                JsonSerializer.Serialize(discoveryRounds));
        if (final is { Error: null, Screen: not null, Print: not null, ScreenToPage: not null }
            && fixture.ExpectedError == null) {
            if (!final.Screen.AsSpan().StartsWith(new byte[] { 137, 80, 78, 71 }) ||
                !final.Print.AsSpan().StartsWith("%PDF-"u8) || !final.ScreenToPage.AsSpan().StartsWith("%PDF-"u8))
                throw new IOException("The isolated renderer returned invalid PNG or PDF signatures.");
            if (final.CaptureUrl != page.DocumentUrl.AbsoluteUri ||
                final.CaptureManifest is not { Length: 71 } manifest || !manifest.StartsWith("sha256:", StringComparison.Ordinal))
                throw new IOException("The isolated renderer returned the wrong capture identity.");
            if (fixture.ExpectedVisibleText != null &&
                (!PdfReadDocument.Open(final.Print).ExtractText().Contains(fixture.ExpectedVisibleText, StringComparison.Ordinal) ||
                 !PdfReadDocument.Open(final.ScreenToPage).ExtractText().Contains(fixture.ExpectedVisibleText, StringComparison.Ordinal)))
                throw new IOException("The rendered PDFs do not contain the script-produced visible text.");
            if (fixture.ExpectBlueInk && (!OfficePngReader.TryDecode(final.Screen, out OfficeRasterImage? raster) || raster == null ||
                !ContainsBlueInk(raster.GetPixels())))
                throw new IOException("The rendered PNG does not contain the styled visible text.");
            if (fixture.ExpectMagentaArea && (!OfficePngReader.TryDecode(final.Screen, out OfficeRasterImage? frameRaster) || frameRaster == null ||
                !ContainsMagentaArea(frameRaster.GetPixels())))
                throw new IOException("The rendered PNG does not contain the child-frame stylesheet background.");
            string caseDirectory = Path.Combine(outputDirectory, fixture.Name);
            Directory.CreateDirectory(caseDirectory);
            await File.WriteAllBytesAsync(Path.Combine(caseDirectory, "screen.png"), final.Screen);
            await File.WriteAllBytesAsync(Path.Combine(caseDirectory, "print.pdf"), final.Print);
            await File.WriteAllBytesAsync(Path.Combine(caseDirectory, "screen-to-page.pdf"), final.ScreenToPage);
        }
    } catch (Exception error) {
        unexpectedException = true;
        errorKind = error.GetType().Name;
        errorMessage = error.Message;
    } finally {
        if (lease != null) {
            try {
                await lease.DisposeAsync();
                removed = true;
            } catch (Exception error) {
                cleanupError = error.GetType().Name + ": " + error.Message;
            }
        }
    }
    bool passed = removed && !unexpectedException && (fixture.ExpectedError == null
        ? final is { Error: null, Screen: not null, Print: not null, ScreenToPage: not null }
        : workerReportedError && errorKind == fixture.ExpectedErrorKind && errorMessage == fixture.ExpectedError);
    return new ProbeResult(fixture.Name, passed, containerName, removed, stopwatch.ElapsedMilliseconds,
        errorKind, errorMessage, cleanupError, discoveryRounds.ToArray(), final?.CaptureManifest,
        final?.Screen == null ? null : Digest(final.Screen),
        final?.Print == null ? null : Digest(final.Print),
        final?.ScreenToPage == null ? null : Digest(final.ScreenToPage));
}

void VerifyIdentity(HtmlPublicRenderResponse response) {
    if (!response.RendererSha256.Equals(rendererSha256, StringComparison.OrdinalIgnoreCase) ||
        !response.WorkerSha256.Equals(workerSha256, StringComparison.OrdinalIgnoreCase) ||
        !response.RendererFilesSha256.Equals(rendererFilesSha256, StringComparison.OrdinalIgnoreCase) ||
        !response.WorkerFilesSha256.Equals(workerFilesSha256, StringComparison.OrdinalIgnoreCase) ||
        !HtmlPublicFontPackage.Matches(fontPackage, response.FontPackage))
        throw new IOException("The isolated renderer or script worker does not match the published files.");
}

static string Digest(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));

static bool DiscoveryRoundsEqual(IReadOnlyList<string[]> actual, IReadOnlyList<string[]> expected) {
    if (actual.Count != expected.Count) return false;
    for (int index = 0; index < actual.Count; index++)
        if (!actual[index].SequenceEqual(expected[index], StringComparer.Ordinal)) return false;
    return true;
}

static bool ContainsBlueInk(byte[] rgba) {
    for (int offset = 0; offset < rgba.Length; offset += 4)
        if (rgba[offset] < 80 && rgba[offset + 1] is > 40 and < 150 && rgba[offset + 2] > 130 && rgba[offset + 3] > 200)
            return true;
    return false;
}

static bool ContainsMagentaArea(byte[] rgba) {
    int pixels = 0;
    for (int offset = 0; offset < rgba.Length; offset += 4) {
        if (rgba[offset] > 160 && rgba[offset + 1] < 80 && rgba[offset + 2] > 80 && rgba[offset + 3] > 200
            && ++pixels >= 100) return true;
    }
    return false;
}

static async Task<string> DrainErrorAsync(Stream stream, CancellationToken token) {
    using var prefix = new MemoryStream();
    var buffer = new byte[4096];
    int read;
    while ((read = await stream.ReadAsync(buffer, token)) > 0)
        if (prefix.Length < 4096) prefix.Write(buffer, 0, (int)Math.Min(read, 4096 - prefix.Length));
    return Encoding.UTF8.GetString(prefix.ToArray());
}

internal sealed record ProbeResource(string Content, string ContentType);
internal sealed record ProbeDynamicResource(string Content, string ContentType, string? ExpectedBody = null,
    IReadOnlyDictionary<string, string>? ExpectedHeaders = null, IReadOnlyList<HtmlRuntimeFetchHop>? Hops = null);
internal sealed record ProbeCase(string Name, string Html, string ReadyExpression, int MaxOutputCharacters,
    string? ExpectedErrorKind = null, string? ExpectedError = null, string? ExpectedVisibleText = null,
    bool ExpectBlueInk = false, bool ExpectMagentaArea = false, IReadOnlyDictionary<string, ProbeResource>? Resources = null,
    string[][]? ExpectedDiscoveryRounds = null, Uri? DocumentUrl = null, double DevicePixelRatio = 1D,
    IReadOnlyDictionary<string, ProbeDynamicResource>? DynamicResponses = null, Uri[]? AllowedOrigins = null,
    long? MaxOutputBytesPerArtifact = null);
internal sealed record ProbeResult(string Name, bool Passed, string? ContainerName, bool ContainerRemoved,
    long ElapsedMilliseconds, string? ErrorKind, string? Error, string? CleanupError,
    string[][] DiscoveryRounds, string? CaptureManifest, string? ScreenSha256, string? PrintSha256,
    string? ScreenToPageSha256);
