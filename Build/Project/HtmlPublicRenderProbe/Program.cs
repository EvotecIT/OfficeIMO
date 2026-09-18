using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Pdf;

if (args.Length != 4)
    throw new ArgumentException("Usage: OfficeIMO.Html.PublicRenderProbe <full-sha256-image-id> <published-renderer-dll> <published-worker-dll> <new-output-directory>");

string imageId = args[0];
string rendererPath = Path.GetFullPath(args[1]);
string workerPath = Path.GetFullPath(args[2]);
string outputDirectory = Path.GetFullPath(args[3]);
if (Directory.Exists(outputDirectory) || File.Exists(outputDirectory))
    throw new IOException("The probe output directory must be new.");
Directory.CreateDirectory(outputDirectory);

string rendererSha256 = Digest(await File.ReadAllBytesAsync(rendererPath));
string workerSha256 = Digest(await File.ReadAllBytesAsync(workerPath));
string rendererFilesSha256 = HtmlPublicArtifactDigest.DirectorySha256(Path.GetDirectoryName(rendererPath)!);
string workerFilesSha256 = HtmlPublicArtifactDigest.DirectorySha256(Path.GetDirectoryName(workerPath)!);
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
    new ProbeCase("dynamic-xhr-headered-get-blocked", """
        <!doctype html><style>body{font:16px sans-serif}#result{color:#0055aa}</style>
        <p id="result">Loading header-varying XHR</p>
        <script>
          const request = new XMLHttpRequest();
          request.open('GET', '/api/header-varying.json');
          request.setRequestHeader('X-Variant', 'private');
          request.onload = () => document.querySelector('#result').textContent = 'Headered XHR was replayed';
          request.onerror = () => document.querySelector('#result').textContent = 'Headered XHR remained offline';
          request.onloadend = () => document.body.dataset.xhrSettled = 'yes';
          request.send();
        </script>
        """, "document.body.dataset.xhrSettled === 'yes'", 8 * 1024 * 1024,
        ExpectedVisibleText: "Headered XHR remained offline", ExpectBlueInk: true,
        ExpectedDiscoveryRounds: []),
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
    new ProbeCase("resource-fanout", "<!doctype html>" + string.Concat(Enumerable.Range(0, 129)
        .Select(index => $"<script src='/asset-{index}.js'></script>")), "true", 8 * 1024 * 1024,
        ExpectedErrorKind: "HtmlScriptRuntimeException",
        ExpectedError: "The document exceeds the pilot's static resource discovery limit."),
    new ProbeCase("capture-output-budget", "<p>" + new string('X', 4096) + "</p>", "true", 1024,
        ExpectedErrorKind: "HtmlScriptRuntimeException", ExpectedError: "Captured data budget exceeded."),
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

await File.WriteAllTextAsync(Path.Combine(outputDirectory, "summary.json"), JsonSerializer.Serialize(new {
    imageId, rendererSha256, workerSha256, rendererFilesSha256, workerFilesSha256,
    capturedAtUtc = DateTimeOffset.UtcNow, isolationPolicy =
        "rootless-podman;seccomp;cgroups-cpu-memory-pids;network-none;read-only;uid-65532;cap-drop-all;no-new-privileges;no-mounts",
    acquisitionCases = acquisition.Results,
    cases = results
}, new JsonSerializerOptions { WriteIndented = true }));
if (acquisition.Results.Any(result => !result.Passed) || results.Any(result => !result.Passed))
    throw new InvalidOperationException("The isolated renderer probe failed; inspect summary.json.");
Console.WriteLine("Isolated renderer probe passed: " + results.Count + " render cases and " +
    acquisition.Results.Count + " acquisition cases.");

async Task<ProbeResult> RunCaseAsync(ProbeCase fixture) {
    using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(60));
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
            ResourcePolicy = new HtmlRuntimeResourcePolicy { AllowNetwork = false, MaxRequests = 64 },
            Timeout = TimeSpan.FromSeconds(3), SessionTimeout = TimeSpan.FromSeconds(12),
            MaxOutputCharacters = fixture.MaxOutputCharacters,
            DevicePixelRatio = fixture.DevicePixelRatio,
            ViewportWidth = 816D,
            ViewportHeight = 720D
        };
        await HtmlRuntimeProtocol.WriteAsync(process.StandardInput.BaseStream,
            new HtmlPublicRenderRequest { Page = page }, 24 * 1024 * 1024, deadline.Token);
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
            discoveryRounds.Add(discoveryRound.ToArray());
            await HtmlRuntimeProtocol.WriteAsync(process.StandardInput.BaseStream,
                new HtmlPublicResourceBatch { Resources = batch.ToArray() }, 24 * 1024 * 1024, deadline.Token);
        }
        process.StandardInput.Close();
        await process.WaitForExitAsync(deadline.Token);
        string stderrText = await stderr;
        if (process.ExitCode != 0) throw new IOException("The isolated renderer exited " + process.ExitCode + ": " + stderrText);
        if (fixture.ExpectedDiscoveryRounds != null &&
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
        !response.WorkerFilesSha256.Equals(workerFilesSha256, StringComparison.OrdinalIgnoreCase))
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
internal sealed record ProbeCase(string Name, string Html, string ReadyExpression, int MaxOutputCharacters,
    string? ExpectedErrorKind = null, string? ExpectedError = null, string? ExpectedVisibleText = null,
    bool ExpectBlueInk = false, bool ExpectMagentaArea = false, IReadOnlyDictionary<string, ProbeResource>? Resources = null,
    string[][]? ExpectedDiscoveryRounds = null, Uri? DocumentUrl = null, double DevicePixelRatio = 1D);
internal sealed record ProbeResult(string Name, bool Passed, string? ContainerName, bool ContainerRemoved,
    long ElapsedMilliseconds, string? ErrorKind, string? Error, string? CleanupError,
    string[][] DiscoveryRounds, string? CaptureManifest, string? ScreenSha256, string? PrintSha256,
    string? ScreenToPageSha256);
