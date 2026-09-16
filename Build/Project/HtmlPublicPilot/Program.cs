using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Html.Runtime;

if (args.Length < 5) throw new ArgumentException("Usage: OfficeIMO.Html.PublicPilot <http(s)-url> <new-output-directory> <full-sha256-image-id> <published-renderer-dll> <published-worker-dll> [--resource=URL] [--host=DNS-name]");
Uri url = HtmlPublicResourceBroker.ValidateUrl(new Uri(args[0], UriKind.Absolute));
string outputDirectory = Path.GetFullPath(args[1]);
if (Directory.Exists(outputDirectory) || File.Exists(outputDirectory))
    throw new IOException("The pilot output directory must not already exist.");
string imageId = args[2];
string rendererPath = Path.GetFullPath(args[3]);
string workerPath = Path.GetFullPath(args[4]);
Uri[] resourceUrls = args.Skip(5).Where(value => value.StartsWith("--resource=", StringComparison.Ordinal))
    .Select(value => HtmlPublicResourceBroker.ValidateUrl(new Uri(value[11..], UriKind.Absolute))).ToArray();
string[] allowedHosts = args.Skip(5).Where(value => value.StartsWith("--host=", StringComparison.Ordinal))
    .Select(value => value[7..]).ToArray();
if (args.Skip(5).Count() != resourceUrls.Length + allowedHosts.Length || resourceUrls.Length > 24)
    throw new ArgumentException("The pilot accepts at most 24 explicit resources and --host entries.");
var broker = new HtmlPublicResourceBroker(new[] { url.IdnHost }.Concat(resourceUrls.Select(resource => resource.IdnHost)).Concat(allowedHosts));
using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(2));
var assets = new List<HtmlPublicResourceResult>();
var skipped = new List<SkippedPublicResource>();
HtmlPublicResourceResult? document = null;
string? containerName = null;
bool removed = false;
HtmlPublicRenderResponse? response = null;
var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
string phase = "acquisition";
Directory.CreateDirectory(outputDirectory);
try {
    document = await broker.FetchAsync(url, deadline.Token);
    string html = HtmlPublicResourceBroker.DecodeUtf8Html(document.Resource);
    foreach (Uri resourceUrl in resourceUrls.DistinctBy(value => HtmlRuntimeResourcePolicy.Key(value), StringComparer.Ordinal))
        assets.Add(await broker.FetchAsync(resourceUrl, deadline.Token));
    int explicitAssetCount = assets.Count;
    static object Provenance(HtmlPublicResourceResult result) => new {
        url = result.Resource.Url.AbsoluteUri,
        finalUrl = result.Resource.FinalUrl.AbsoluteUri,
        result.Resource.StatusCode,
        result.Resource.ContentType,
        bytes = result.Resource.Length,
        result.Sha256,
        result.FetchedAtUtc,
        connectedAddress = result.ConnectedAddress.ToString(),
        redirects = result.Redirects.Select(hop => new {
            from = hop.From.AbsoluteUri, to = hop.To.AbsoluteUri, hop.StatusCode,
            connectedAddress = hop.ConnectedAddress.ToString()
        }).ToArray()
    };
    async Task PersistAcquisitionAsync() {
        await File.WriteAllBytesAsync(Path.Combine(outputDirectory, "response.html"), document.Resource.Content);
        if (assets.Count != 0) Directory.CreateDirectory(Path.Combine(outputDirectory, "assets"));
        foreach (HtmlPublicResourceResult asset in assets) {
            string path = Path.Combine(outputDirectory, "assets", asset.Sha256);
            if (!File.Exists(path)) await File.WriteAllBytesAsync(path, asset.Resource.Content);
        }
        await File.WriteAllTextAsync(Path.Combine(outputDirectory, "acquisition.json"), JsonSerializer.Serialize(
            new { document = Provenance(document), assets = assets.Select(Provenance).ToArray(), skipped }, jsonOptions));
    }
    await PersistAcquisitionAsync();

    phase = "preflight";
    string expectedRenderer = Convert.ToHexStringLower(SHA256.HashData(await File.ReadAllBytesAsync(rendererPath, deadline.Token)));
    string expectedWorker = Convert.ToHexStringLower(SHA256.HashData(await File.ReadAllBytesAsync(workerPath, deadline.Token)));
    string expectedRendererFiles = HtmlPublicArtifactDigest.DirectorySha256(Path.GetDirectoryName(rendererPath)!);
    string expectedWorkerFiles = HtmlPublicArtifactDigest.DirectorySha256(Path.GetDirectoryName(workerPath)!);
    var page = new HtmlScriptRequest {
        Profile = HtmlRuntimeProfile.WebApplicationV1,
        Html = html,
        DocumentUrl = document.Resource.FinalUrl,
        Resources = assets.Select(asset => asset.Resource).ToArray(),
        ResourcePolicy = new HtmlRuntimeResourcePolicy {
            AllowNetwork = false, AllowedOrigins = broker.AllowedOrigins,
            MaxRequests = 64, MaxResourceBytes = 4 * 1024 * 1024, MaxTotalBytes = 16 * 1024 * 1024
        },
        Timeout = TimeSpan.FromSeconds(20), SessionTimeout = TimeSpan.FromMinutes(2)
    }.Snapshot();

    phase = "isolated-render";
    HtmlOciWorkerLease? lease = null;
    try {
        lease = await HtmlOciWorkerLease.StartAsync("podman", imageId, deadline.Token);
        containerName = lease.ContainerName;
        Process process = lease.Process;
        Task<string> stderr = DrainErrorAsync(process.StandardError.BaseStream, deadline.Token);
        await HtmlRuntimeProtocol.WriteAsync(process.StandardInput.BaseStream,
            new HtmlPublicRenderRequest { Page = page }, 24 * 1024 * 1024, deadline.Token);
        var known = new HashSet<string>(assets.Select(asset => HtmlRuntimeResourcePolicy.Key(asset.Resource.Url)), StringComparer.Ordinal) {
            HtmlRuntimeResourcePolicy.Key(document.Resource.Url)
        };
        for (int round = 0; ; round++) {
            phase = "resource-discovery";
            response = await HtmlRuntimeProtocol.ReadAsync<HtmlPublicRenderResponse>(
                process.StandardOutput.BaseStream, 24 * 1024 * 1024, deadline.Token)
                ?? throw new HtmlScriptRuntimeException("The isolated renderer returned no discovery response.");
            VerifyIdentity(response, expectedRenderer, expectedWorker, expectedRendererFiles, expectedWorkerFiles);
            if (response.Error != null) throw new HtmlScriptRuntimeException(response.ErrorKind + ": " + response.Error);
            if (response.DiscoveryComplete) break;
            if (round >= 16 || response.DiscoveryUrls == null || response.DiscoveryUrls.Length is 0 or > 128)
                throw new HtmlScriptRuntimeException("The isolated renderer exceeded the resource discovery limit.");
            var batch = new List<HtmlRuntimeResource>();
            foreach (string candidate in response.DiscoveryUrls) {
                if (!Uri.TryCreate(candidate, UriKind.Absolute, out Uri? resourceUrl)) {
                    skipped.Add(new SkippedPublicResource(candidate, "invalid-url"));
                    continue;
                }
                try { HtmlPublicResourceBroker.ValidateUrl(resourceUrl); }
                catch (ArgumentException) {
                    skipped.Add(new SkippedPublicResource(candidate, "unsupported-url"));
                    continue;
                }
                if (!broker.AllowsHost(resourceUrl)) {
                    skipped.Add(new SkippedPublicResource(candidate, "host-not-allowed"));
                    continue;
                }
                if (!known.Add(HtmlRuntimeResourcePolicy.Key(resourceUrl))) continue;
                if (assets.Count >= 24) {
                    skipped.Add(new SkippedPublicResource(candidate, "resource-count-limit"));
                    continue;
                }
                HtmlPublicResourceResult asset = await broker.FetchAsync(resourceUrl, deadline.Token);
                assets.Add(asset);
                batch.Add(asset.Resource);
                await PersistAcquisitionAsync();
            }
            await PersistAcquisitionAsync();
            await HtmlRuntimeProtocol.WriteAsync(process.StandardInput.BaseStream,
                new HtmlPublicResourceBatch { Resources = batch.ToArray() }, 24 * 1024 * 1024, deadline.Token);
        }
        process.StandardInput.Close();
        phase = "isolated-render";
        response = await HtmlRuntimeProtocol.ReadAsync<HtmlPublicRenderResponse>(
            process.StandardOutput.BaseStream, 24 * 1024 * 1024, deadline.Token)
            ?? throw new HtmlScriptRuntimeException("The isolated renderer returned no final response.");
        await process.WaitForExitAsync(deadline.Token);
        string errorOutput = await stderr;
        if (process.ExitCode != 0)
            throw new HtmlScriptRuntimeException("The isolated renderer exited " + process.ExitCode + ": " + errorOutput);
        VerifyIdentity(response, expectedRenderer, expectedWorker, expectedRendererFiles, expectedWorkerFiles);
    } finally {
        if (lease != null) {
            await lease.DisposeAsync();
            removed = true;
        }
    }
    if (response.Error != null) throw new HtmlScriptRuntimeException(response.ErrorKind + ": " + response.Error);
    byte[] screen = response.Screen ?? throw new HtmlScriptRuntimeException("The isolated screen output is missing.");
    byte[] print = response.Print ?? throw new HtmlScriptRuntimeException("The isolated print output is missing.");
    byte[] screenToPage = response.ScreenToPage ?? throw new HtmlScriptRuntimeException("The isolated screen-to-page output is missing.");
    phase = "output";
    await File.WriteAllBytesAsync(Path.Combine(outputDirectory, "screen.png"), screen);
    await File.WriteAllBytesAsync(Path.Combine(outputDirectory, "print.pdf"), print);
    await File.WriteAllBytesAsync(Path.Combine(outputDirectory, "screen-to-page.pdf"), screenToPage);
    var outcome = new {
        provider = "officeimo.oci-public-pilot", innerProvider = response.ProviderId,
        response.CaptureUrl, response.CaptureManifest, response.TraceEntries,
        imageId, containerName, containerRemoved = removed,
        isolationPolicy = "rootless-podman;seccomp;cgroups-cpu-memory-pids;network-none;read-only;uid-65532;cap-drop-all;no-new-privileges;no-mounts",
        response.RendererSha256, response.WorkerSha256,
        response.RendererFilesSha256, response.WorkerFilesSha256,
        discoveredAssets = assets.Count - explicitAssetCount, skippedResources = skipped,
        outputs = new[] {
            Output("screen.png", screen), Output("print.pdf", print), Output("screen-to-page.pdf", screenToPage)
        }
    };
    await File.WriteAllTextAsync(Path.Combine(outputDirectory, "outcome.json"), JsonSerializer.Serialize(outcome, jsonOptions));
    Console.WriteLine("Isolated public-page output captured for " + response.CaptureUrl);
} catch (Exception error) {
    await File.WriteAllTextAsync(Path.Combine(outputDirectory, "failure.json"), JsonSerializer.Serialize(new {
        errorKind = error.GetType().Name, error = error.Message, phase,
        requestedUrl = url.AbsoluteUri, imageId, containerName,
        containerRemoved = removed, rendererSha256 = response?.RendererSha256,
        workerSha256 = response?.WorkerSha256, rendererFilesSha256 = response?.RendererFilesSha256,
        workerFilesSha256 = response?.WorkerFilesSha256, trace = response?.TraceEntries,
        documentSha256 = document?.Sha256, assets = assets.Count, skippedResources = skipped
    }, jsonOptions));
    Console.Error.WriteLine("Pilot failed: " + error.GetType().Name + ": " + error.Message);
    Environment.ExitCode = 2;
}

static object Output(string file, byte[] bytes) => new {
    file, bytes = bytes.Length, sha256 = Convert.ToHexStringLower(SHA256.HashData(bytes))
};

static void VerifyIdentity(HtmlPublicRenderResponse response, string renderer, string worker,
    string rendererFiles, string workerFiles) {
    if (!response.RendererSha256.Equals(renderer, StringComparison.OrdinalIgnoreCase) ||
        !response.WorkerSha256.Equals(worker, StringComparison.OrdinalIgnoreCase) ||
        !response.RendererFilesSha256.Equals(rendererFiles, StringComparison.OrdinalIgnoreCase) ||
        !response.WorkerFilesSha256.Equals(workerFiles, StringComparison.OrdinalIgnoreCase))
        throw new HtmlScriptRuntimeException("The isolated image does not contain the expected renderer and script worker builds.");
}

static async Task<string> DrainErrorAsync(Stream stream, CancellationToken token) {
    var prefix = new MemoryStream();
    var buffer = new byte[4096];
    int read;
    while ((read = await stream.ReadAsync(buffer, token)) != 0) {
        if (prefix.Length < 2048) prefix.Write(buffer, 0, (int)Math.Min(read, 2048 - prefix.Length));
    }
    return Encoding.UTF8.GetString(prefix.ToArray());
}

internal sealed record SkippedPublicResource(string Url, string Reason);
