using System.Text.Json;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Rendering;

if (args.Length < 6) throw new ArgumentException(
    "Usage: OfficeIMO.Html.PublicPilot <http(s)-url> <new-output-directory> <full-sha256-image-id> <published-renderer-dll> <published-worker-dll> --license=TEXT [--scenario=ID] [--resource=URL] [--host=DNS-name] [--retain-input] [--podman-command=PATH] [--podman-arg=VALUE] [--timeout-seconds=SECONDS]");

Uri url = new(args[0], UriKind.Absolute);
string outputDirectory = Path.GetFullPath(args[1]);
if (Directory.Exists(outputDirectory) || File.Exists(outputDirectory))
    throw new IOException("The pilot output directory must not already exist.");
string? license = Value("--license=");
if (string.IsNullOrWhiteSpace(license)) throw new ArgumentException("The pilot requires --license=TEXT.");
string scenario = Value("--scenario=") ?? "public-page";
string podmanCommand = Value("--podman-command=") ?? "podman";
string[] podmanArguments = Values("--podman-arg=");
TimeSpan timeout = TimeSpan.FromSeconds(ParseTimeout(Value("--timeout-seconds=")));
Uri[] resources = Values("--resource=").Select(value => new Uri(value, UriKind.Absolute)).ToArray();
string[] hosts = Values("--host=");
bool retainInput = args.Contains("--retain-input", StringComparer.Ordinal);
string[] knownPrefixes = ["--license=", "--scenario=", "--podman-command=", "--podman-arg=", "--resource=", "--host=", "--timeout-seconds="];
if (args.Skip(5).Any(value => value != "--retain-input" &&
        !knownPrefixes.Any(prefix => value.StartsWith(prefix, StringComparison.Ordinal))))
    throw new ArgumentException("The pilot received an unknown option.");

Directory.CreateDirectory(outputDirectory);
var json = new JsonSerializerOptions { WriteIndented = true };
try {
    HtmlIsolatedPublicPageResult result = await HtmlIsolatedPublicPageWorkflow.RunAsync(
        new HtmlIsolatedPublicPageExecutionOptions {
            ImageId = args[2],
            PodmanCommand = podmanCommand,
            PodmanCommandArguments = podmanArguments,
            PublishedRendererAssemblyPath = args[3],
            PublishedWorkerAssemblyPath = args[4],
            OperationTimeout = timeout
        },
        new HtmlIsolatedPublicPageRequest {
            ScenarioId = scenario,
            Url = url,
            SourceLicense = license,
            SeedResourceUrls = resources,
            AllowedHosts = hosts,
            RetainInputBytes = retainInput
        });

    foreach (HtmlIsolatedPageOutput output in result.Outputs)
        await File.WriteAllBytesAsync(Path.Combine(outputDirectory, output.Name), output.Content);
    await PersistAcquisitionAsync(result.Resources, result.SkippedResources);
    await File.WriteAllTextAsync(Path.Combine(outputDirectory, "outcome.json"), JsonSerializer.Serialize(new {
        result.ScenarioId,
        result.SourceLicense,
        result.RequestedUrl,
        result.ProviderId,
        result.CaptureUrl,
        result.CaptureManifest,
        result.TraceEntries,
        result.ImageId,
        result.ContainerName,
        result.ContainerRemoved,
        result.IsolationPolicy,
        result.IsolationProfile,
        result.UnsupportedFeatures,
        result.RendererSha256,
        result.WorkerSha256,
        result.RendererFilesSha256,
        result.WorkerFilesSha256,
        outputs = result.Outputs.Select(output => new {
            file = output.Name, bytes = output.Content.Length, output.Sha256, output.MediaType
        }).ToArray()
    }, json));
    Console.WriteLine("Isolated public-page output captured for " + result.CaptureUrl);
} catch (HtmlIsolatedPublicPageCanceledException error) {
    await PersistFailureAsync(error, error.Evidence);
} catch (HtmlIsolatedPublicPageException error) {
    await PersistFailureAsync(error, error.Evidence);
} catch (Exception error) {
    await PersistFailureAsync(error, null);
}

string? Value(string prefix) => Values(prefix).SingleOrDefault();
string[] Values(string prefix) => args.Skip(5).Where(value => value.StartsWith(prefix, StringComparison.Ordinal))
    .Select(value => value[prefix.Length..]).ToArray();

static int ParseTimeout(string? value) {
    if (value == null) return 120;
    if (!int.TryParse(value, System.Globalization.NumberStyles.None,
            System.Globalization.CultureInfo.InvariantCulture, out int seconds) || seconds is < 1 or > 600)
        throw new ArgumentException("--timeout-seconds must be an integer from 1 through 600.");
    return seconds;
}

async Task PersistAcquisitionAsync(IReadOnlyList<HtmlPublicResourceEvidence> acquired,
    IReadOnlyList<HtmlPublicSkippedResource> omitted) {
    if (retainInput) {
        Directory.CreateDirectory(Path.Combine(outputDirectory, "inputs"));
        foreach (HtmlPublicResourceEvidence resource in acquired.Where(resource => resource.Content != null))
            await File.WriteAllBytesAsync(Path.Combine(outputDirectory, "inputs", resource.Sha256), resource.Content!.Value);
    }
    await File.WriteAllTextAsync(Path.Combine(outputDirectory, "acquisition.json"), JsonSerializer.Serialize(new {
        ScenarioId = scenario,
        SourceLicense = license,
        retainedInputBytes = retainInput,
        resources = acquired.Select(ResourceEvidence).ToArray(),
        SkippedResources = omitted
    }, json));
}

async Task PersistFailureAsync(Exception error, HtmlIsolatedPublicPageFailureEvidence? evidence) {
    if (evidence != null) await PersistAcquisitionAsync(evidence.Resources, evidence.SkippedResources);
    await File.WriteAllTextAsync(Path.Combine(outputDirectory, "failure.json"), JsonSerializer.Serialize(new {
        errorKind = error.GetType().Name,
        error = error.Message,
        phase = evidence?.Phase.ToString() ?? HtmlIsolatedPublicPagePhase.Admission.ToString(),
        requestedUrl = url.AbsoluteUri,
        imageId = args[2],
        scenario,
        sourceLicense = license,
        operationTimeoutSeconds = timeout.TotalSeconds,
        maximumCleanupSeconds = HtmlIsolatedPublicPageExecutionOptions.MaximumCleanupDuration.TotalSeconds,
        containerName = evidence?.ContainerName,
        containerRemoved = evidence?.ContainerRemoved,
        cleanupError = evidence?.CleanupError,
        providerId = evidence?.ProviderId,
        captureUrl = evidence?.CaptureUrl,
        trace = evidence?.TraceEntries,
        expectedRendererSha256 = evidence?.ExpectedRendererSha256,
        expectedWorkerSha256 = evidence?.ExpectedWorkerSha256,
        expectedRendererFilesSha256 = evidence?.ExpectedRendererFilesSha256,
        expectedWorkerFilesSha256 = evidence?.ExpectedWorkerFilesSha256,
        reportedRendererSha256 = evidence?.ReportedRendererSha256,
        reportedWorkerSha256 = evidence?.ReportedWorkerSha256,
        reportedRendererFilesSha256 = evidence?.ReportedRendererFilesSha256,
        reportedWorkerFilesSha256 = evidence?.ReportedWorkerFilesSha256,
        acquiredResources = evidence?.Resources.Count ?? 0,
        skippedResources = evidence?.SkippedResources ?? Array.Empty<HtmlPublicSkippedResource>()
    }, json));
    Console.Error.WriteLine("Pilot failed: " + error.GetType().Name + ": " + error.Message);
    Environment.ExitCode = 2;
}

static object ResourceEvidence(HtmlPublicResourceEvidence resource) => new {
    url = resource.Url.AbsoluteUri,
    finalUrl = resource.FinalUrl.AbsoluteUri,
    resource.StatusCode,
    resource.ContentType,
    resource.ByteCount,
    resource.Sha256,
    resource.FetchedAtUtc,
    connectedAddress = resource.ConnectedAddress.ToString(),
    redirects = resource.Redirects.Select(redirect => new {
        from = redirect.From.AbsoluteUri,
        to = redirect.To.AbsoluteUri,
        redirect.StatusCode,
        connectedAddress = redirect.ConnectedAddress.ToString()
    }).ToArray()
};
