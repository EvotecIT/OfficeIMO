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
var cases = new[] {
    new ProbeCase("malformed-markup", """
        <!doctype html><style>body{font:16px sans-serif}p{color:#0055aa}</style>
        <main><table><tr><td><p id=result>Before script
        <script>document.querySelector('#result').textContent='Recovered by script';</script>
        """, "document.querySelector('#result')?.textContent === 'Recovered by script'", 8 * 1024 * 1024, null, null),
    new ProbeCase("resource-fanout", "<!doctype html>" + string.Concat(Enumerable.Range(0, 129)
        .Select(index => $"<script src='/asset-{index}.js'></script>")), "true", 8 * 1024 * 1024,
        "HtmlScriptRuntimeException", "The document exceeds the pilot's static resource discovery limit."),
    new ProbeCase("capture-output-budget", "<p>" + new string('X', 4096) + "</p>", "true", 1024,
        "HtmlScriptRuntimeException", "Captured data budget exceeded."),
    new ProbeCase("runaway-script", "<script>while(true){}</script>", "true", 8 * 1024 * 1024,
        "TimeoutException", "The runtime command exceeded its deadline.")
};
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
    cases = results
}, new JsonSerializerOptions { WriteIndented = true }));
if (results.Any(result => !result.Passed)) throw new InvalidOperationException("The isolated renderer hostile-input probe failed; inspect summary.json.");
Console.WriteLine("Isolated renderer hostile-input probe passed: " + results.Count + " cases.");

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
    var stopwatch = Stopwatch.StartNew();
    try {
        lease = await HtmlOciWorkerLease.StartAsync("podman", imageId, deadline.Token);
        containerName = lease.ContainerName;
        Process process = lease.Process;
        Task<string> stderr = DrainErrorAsync(process.StandardError.BaseStream, deadline.Token);
        var page = new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            Html = fixture.Html,
            DocumentUrl = new Uri("https://fixture.officeimo.invalid/"),
            ReadyExpression = fixture.ReadyExpression,
            ResourcePolicy = new HtmlRuntimeResourcePolicy { AllowNetwork = false, MaxRequests = 64 },
            Timeout = TimeSpan.FromSeconds(3), SessionTimeout = TimeSpan.FromSeconds(12),
            MaxOutputCharacters = fixture.MaxOutputCharacters
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
            await HtmlRuntimeProtocol.WriteAsync(process.StandardInput.BaseStream,
                new HtmlPublicResourceBatch(), 24 * 1024 * 1024, deadline.Token);
        }
        process.StandardInput.Close();
        await process.WaitForExitAsync(deadline.Token);
        string stderrText = await stderr;
        if (process.ExitCode != 0) throw new IOException("The isolated renderer exited " + process.ExitCode + ": " + stderrText);
        if (final is { Error: null, Screen: not null, Print: not null, ScreenToPage: not null }
            && fixture.ExpectedError == null) {
            if (!final.Screen.AsSpan().StartsWith(new byte[] { 137, 80, 78, 71 }) ||
                !final.Print.AsSpan().StartsWith("%PDF-"u8) || !final.ScreenToPage.AsSpan().StartsWith("%PDF-"u8))
                throw new IOException("The isolated renderer returned invalid PNG or PDF signatures.");
            if (final.CaptureUrl != page.DocumentUrl.AbsoluteUri ||
                final.CaptureManifest is not { Length: 71 } manifest || !manifest.StartsWith("sha256:", StringComparison.Ordinal))
                throw new IOException("The isolated renderer returned the wrong capture identity.");
            const string expectedText = "Recovered by script";
            if (!PdfReadDocument.Open(final.Print).ExtractText().Contains(expectedText, StringComparison.Ordinal) ||
                !PdfReadDocument.Open(final.ScreenToPage).ExtractText().Contains(expectedText, StringComparison.Ordinal))
                throw new IOException("The rendered PDFs do not contain the script-produced visible text.");
            if (!OfficePngReader.TryDecode(final.Screen, out OfficeRasterImage? raster) || raster == null ||
                !ContainsBlueInk(raster.GetPixels()))
                throw new IOException("The rendered PNG does not contain the styled visible text.");
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
        errorKind, errorMessage, cleanupError, final?.CaptureManifest,
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

static bool ContainsBlueInk(byte[] rgba) {
    for (int offset = 0; offset < rgba.Length; offset += 4)
        if (rgba[offset] < 80 && rgba[offset + 1] is > 40 and < 150 && rgba[offset + 2] > 130 && rgba[offset + 3] > 200)
            return true;
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

internal sealed record ProbeCase(string Name, string Html, string ReadyExpression,
    int MaxOutputCharacters, string? ExpectedErrorKind, string? ExpectedError);
internal sealed record ProbeResult(string Name, bool Passed, string? ContainerName, bool ContainerRemoved,
    long ElapsedMilliseconds, string? ErrorKind, string? Error, string? CleanupError, string? CaptureManifest,
    string? ScreenSha256, string? PrintSha256, string? ScreenToPageSha256);
