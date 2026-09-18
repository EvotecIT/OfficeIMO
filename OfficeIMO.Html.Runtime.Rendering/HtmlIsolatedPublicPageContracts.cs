using System.Net;
using OfficeIMO.Html.Runtime;

namespace OfficeIMO.Html.Runtime.Rendering;

/// <summary>OS-enforced execution boundary used for untrusted public-page content.</summary>
public enum HtmlPublicPageIsolationProfile {
    /// <summary>Rootless Linux OCI container with no network, read-only storage, seccomp and cgroup limits.</summary>
    NetworklessRootlessOciV1
}

/// <summary>Last workflow phase reached before an isolated public-page failure.</summary>
public enum HtmlIsolatedPublicPagePhase {
    /// <summary>Request and execution options are being validated.</summary>
    Admission,
    /// <summary>The bounded host broker is acquiring public bytes.</summary>
    Acquisition,
    /// <summary>Published payload identities are being calculated.</summary>
    Preflight,
    /// <summary>The isolated container is being created and inspected.</summary>
    IsolatedStartup,
    /// <summary>The isolated worker is requesting broker-authorized resources.</summary>
    ResourceDiscovery,
    /// <summary>The isolated worker is scripting, capturing, or rendering the page.</summary>
    Rendering,
    /// <summary>The isolated container is being stopped and removed.</summary>
    Cleanup,
    /// <summary>Encoded outputs are being validated and returned.</summary>
    Output
}

/// <summary>Host acquisition and runtime settings for one named public-page run.</summary>
public sealed class HtmlIsolatedPublicPageRequest {
    /// <summary>Stable scenario identity used in retained evidence.</summary>
    public string ScenarioId { get; init; } = string.Empty;
    /// <summary>Public HTTP(S) document URL acquired by the bounded host broker.</summary>
    public Uri Url { get; init; } = new("https://officeimo.invalid/");
    /// <summary>License or usage-rights note retained with the run.</summary>
    public string SourceLicense { get; init; } = string.Empty;
    /// <summary>Optional resources fetched before isolated discovery begins.</summary>
    public IReadOnlyList<Uri> SeedResourceUrls { get; init; } = Array.Empty<Uri>();
    /// <summary>Additional DNS hosts authorized for redirects and discovered resources.</summary>
    public IReadOnlyList<string> AllowedHosts { get; init; } = Array.Empty<string>();
    /// <summary>Runtime limits and readiness. HTML, URL, scripts, resources, and network authority are supplied by this workflow.</summary>
    public HtmlScriptRequest Runtime { get; init; } = new() {
        Profile = HtmlRuntimeProfile.WebApplicationV1,
        ViewportWidth = 816D,
        ViewportHeight = 720D,
        Timeout = TimeSpan.FromSeconds(20),
        SessionTimeout = TimeSpan.FromMinutes(2)
    };
    /// <summary>Retains acquired response bytes in the result. The default retains only provenance and digests.</summary>
    public bool RetainInputBytes { get; init; }

    internal Snapshot Validate() {
        if (string.IsNullOrWhiteSpace(ScenarioId) || ScenarioId.Length > 128)
            throw new ArgumentException("ScenarioId is required and cannot exceed 128 characters.", nameof(ScenarioId));
        if (ScenarioId.Any(character => !(char.IsAsciiLetterOrDigit(character) || character is '-' or '_' or '.')))
            throw new ArgumentException("ScenarioId accepts ASCII letters, digits, '.', '-', and '_' only.", nameof(ScenarioId));
        if (string.IsNullOrWhiteSpace(SourceLicense) || SourceLicense.Length > 512)
            throw new ArgumentException("SourceLicense is required and cannot exceed 512 characters.", nameof(SourceLicense));
        Uri url = HtmlPublicResourceBroker.ValidateUrl(Url ?? throw new ArgumentNullException(nameof(Url)));
        ArgumentNullException.ThrowIfNull(SeedResourceUrls);
        ArgumentNullException.ThrowIfNull(AllowedHosts);
        if (SeedResourceUrls.Count > 24) throw new ArgumentException("At most 24 seed resources are allowed.", nameof(SeedResourceUrls));
        if (AllowedHosts.Count > 15) throw new ArgumentException("At most 15 additional hosts are allowed.", nameof(AllowedHosts));
        Uri[] resources = SeedResourceUrls.Select(resource => HtmlPublicResourceBroker.ValidateUrl(
            resource ?? throw new ArgumentException("A seed resource URL cannot be null.", nameof(SeedResourceUrls)))).ToArray();
        string[] hosts = AllowedHosts.Select(host => host ?? throw new ArgumentException(
            "An allowed host cannot be null.", nameof(AllowedHosts))).ToArray();
        HtmlScriptRequest runtime = (Runtime ?? throw new ArgumentNullException(nameof(Runtime))).Snapshot();
        if (runtime.Profile != HtmlRuntimeProfile.WebApplicationV1)
            throw new ArgumentException("The isolated public-page workflow requires WebApplicationV1.", nameof(Runtime));
        if (runtime.Html.Length != 0 || runtime.Scripts.Count != 0 || runtime.Resources.Count != 0 ||
            runtime.ResourcePolicy.AllowNetwork || runtime.ResourcePolicy.AllowedOrigins.Count != 0)
            throw new ArgumentException("Runtime HTML, scripts, resources, and network access are owned by the isolated public-page workflow.", nameof(Runtime));
        if (runtime.Timeout > TimeSpan.FromSeconds(30) || runtime.SessionTimeout > TimeSpan.FromMinutes(2))
            throw new ArgumentException("Public-page command and session deadlines exceed the isolated profile.", nameof(Runtime));
        return new Snapshot(ScenarioId, url, SourceLicense, resources, hosts, runtime, RetainInputBytes);
    }

    internal sealed record Snapshot(string ScenarioId, Uri Url, string SourceLicense, Uri[] SeedResourceUrls,
        string[] AllowedHosts, HtmlScriptRequest Runtime, bool RetainInputBytes);
}

/// <summary>Immutable OCI image and container-engine command used by the isolated renderer.</summary>
public sealed class HtmlIsolatedPublicPageExecutionOptions {
    /// <summary>Full immutable <c>sha256:</c> image ID.</summary>
    public string ImageId { get; init; } = string.Empty;
    /// <summary>Command executable used to invoke Podman. Use <c>wsl.exe</c> with prefix arguments on Windows.</summary>
    public string PodmanCommand { get; init; } = "podman";
    /// <summary>Arguments inserted before every Podman command, such as <c>-d Ubuntu --exec podman</c>.</summary>
    public IReadOnlyList<string> PodmanCommandArguments { get; init; } = Array.Empty<string>();
    /// <summary>Published renderer assembly whose complete directory must match the image payload.</summary>
    public string PublishedRendererAssemblyPath { get; init; } = string.Empty;
    /// <summary>Published script-worker assembly whose complete directory must match the image payload.</summary>
    public string PublishedWorkerAssemblyPath { get; init; } = string.Empty;
    /// <summary>Acquisition, preflight, discovery and isolated-execution deadline. Verified cleanup has a separate fixed budget.</summary>
    public TimeSpan OperationTimeout { get; init; } = TimeSpan.FromMinutes(2);
    /// <summary>Maximum additional time allowed for fail-safe container removal after the operation ends.</summary>
    public static TimeSpan MaximumCleanupDuration => HtmlOciWorkerLease.MaximumCleanupDuration;

    internal Snapshot Validate() {
        if (string.IsNullOrWhiteSpace(ImageId) || ImageId.Length != 71 || !ImageId.StartsWith("sha256:", StringComparison.Ordinal)
            || ImageId.AsSpan(7).ToString().Any(character => !Uri.IsHexDigit(character)))
            throw new ArgumentException("ImageId must be a full immutable sha256 image ID.", nameof(ImageId));
        if (string.IsNullOrWhiteSpace(PodmanCommand))
            throw new ArgumentException("A Podman command is required.", nameof(PodmanCommand));
        ArgumentNullException.ThrowIfNull(PodmanCommandArguments);
        if (PodmanCommandArguments.Count > 16 || PodmanCommandArguments.Any(argument => string.IsNullOrWhiteSpace(argument) || argument.Length > 1024))
            throw new ArgumentException("Podman command prefix arguments are invalid.", nameof(PodmanCommandArguments));
        string renderer = RequiredFile(PublishedRendererAssemblyPath, nameof(PublishedRendererAssemblyPath));
        string worker = RequiredFile(PublishedWorkerAssemblyPath, nameof(PublishedWorkerAssemblyPath));
        if (OperationTimeout <= TimeSpan.Zero || OperationTimeout > TimeSpan.FromMinutes(10))
            throw new ArgumentOutOfRangeException(nameof(OperationTimeout));
        return new Snapshot(ImageId.ToLowerInvariant(), PodmanCommand, PodmanCommandArguments.ToArray(), renderer, worker, OperationTimeout);
    }

    private static string RequiredFile(string path, string parameter) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A published assembly path is required.", parameter);
        string fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The published assembly was not found.", fullPath);
        return fullPath;
    }

    internal sealed record Snapshot(string ImageId, string Executable, string[] Arguments,
        string RendererPath, string WorkerPath, TimeSpan OperationTimeout);
}

/// <summary>One acquired response and its redirect and connection provenance.</summary>
public sealed class HtmlPublicResourceEvidence {
    internal HtmlPublicResourceEvidence(HtmlPublicResourceResult result, bool retainBytes) {
        Url = result.Resource.Url;
        FinalUrl = result.Resource.FinalUrl;
        StatusCode = result.Resource.StatusCode;
        ContentType = result.Resource.ContentType;
        ByteCount = result.Resource.Length;
        Sha256 = result.Sha256;
        FetchedAtUtc = result.FetchedAtUtc;
        ConnectedAddress = result.ConnectedAddress;
        Redirects = Array.AsReadOnly(result.Redirects.Select(redirect => new HtmlPublicRedirectEvidence(
            redirect.From, redirect.To, redirect.StatusCode, redirect.ConnectedAddress)).ToArray());
        Content = retainBytes ? (ReadOnlyMemory<byte>?)result.Resource.Content.ToArray() : null;
    }

    /// <summary>Originally requested resource URL.</summary>
    public Uri Url { get; }
    /// <summary>Final response URL after validated redirects.</summary>
    public Uri FinalUrl { get; }
    /// <summary>Final HTTP response status.</summary>
    public int StatusCode { get; }
    /// <summary>Validated response media type.</summary>
    public string ContentType { get; }
    /// <summary>Acquired response byte count.</summary>
    public long ByteCount { get; }
    /// <summary>Lowercase SHA-256 digest of the acquired bytes.</summary>
    public string Sha256 { get; }
    /// <summary>UTC time at which acquisition completed.</summary>
    public DateTimeOffset FetchedAtUtc { get; }
    /// <summary>Public IPv4 address used for the final direct connection.</summary>
    public IPAddress ConnectedAddress { get; }
    /// <summary>Validated redirect hops in request order.</summary>
    public IReadOnlyList<HtmlPublicRedirectEvidence> Redirects { get; }
    /// <summary>Acquired bytes when retention was explicitly requested; otherwise <see langword="null"/>.</summary>
    public ReadOnlyMemory<byte>? Content { get; }
}

/// <summary>One retained public HTTP redirect hop.</summary>
public sealed class HtmlPublicRedirectEvidence {
    internal HtmlPublicRedirectEvidence(Uri from, Uri to, int statusCode, IPAddress connectedAddress) {
        From = from;
        To = to;
        StatusCode = statusCode;
        ConnectedAddress = connectedAddress;
    }

    /// <summary>URL which returned the redirect.</summary>
    public Uri From { get; }
    /// <summary>Validated redirect target.</summary>
    public Uri To { get; }
    /// <summary>Redirect HTTP status.</summary>
    public int StatusCode { get; }
    /// <summary>Public IPv4 address used for the redirecting connection.</summary>
    public IPAddress ConnectedAddress { get; }
}

/// <summary>A discovered URL omitted by the public acquisition policy.</summary>
public sealed class HtmlPublicSkippedResource {
    internal HtmlPublicSkippedResource(string url, string reason) {
        Url = url;
        Reason = reason;
    }

    /// <summary>Discovered URL text.</summary>
    public string Url { get; }
    /// <summary>Stable policy reason for omitting the resource.</summary>
    public string Reason { get; }
}

/// <summary>One encoded output returned by the isolated renderer.</summary>
public sealed class HtmlIsolatedPageOutput {
    internal HtmlIsolatedPageOutput(string name, string mediaType, byte[] content) {
        Name = name;
        MediaType = mediaType;
        Content = content.ToArray();
        Sha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(content)).ToLowerInvariant();
    }

    /// <summary>Suggested artifact file name.</summary>
    public string Name { get; }
    /// <summary>Output media type.</summary>
    public string MediaType { get; }
    /// <summary>Encoded artifact bytes.</summary>
    public ReadOnlyMemory<byte> Content { get; }
    /// <summary>Lowercase SHA-256 digest of the encoded bytes.</summary>
    public string Sha256 { get; }
}

/// <summary>Bounded partial provenance and cleanup evidence retained when an isolated public-page run fails.</summary>
public sealed class HtmlIsolatedPublicPageFailureEvidence {
    internal HtmlIsolatedPublicPageFailureEvidence(HtmlIsolatedPublicPagePhase phase,
        IReadOnlyList<HtmlPublicResourceEvidence> resources, IReadOnlyList<HtmlPublicSkippedResource> skippedResources,
        string imageId, string? containerName, bool? containerRemoved, string? cleanupError,
        HtmlPublicRenderResponse? response, string? rendererSha256, string? workerSha256,
        string? rendererFilesSha256, string? workerFilesSha256) {
        Phase = phase;
        Resources = resources;
        SkippedResources = skippedResources;
        ImageId = imageId;
        ContainerName = containerName;
        ContainerRemoved = containerRemoved;
        CleanupError = cleanupError;
        ProviderId = response?.ProviderId;
        CaptureUrl = Uri.TryCreate(response?.CaptureUrl, UriKind.Absolute, out Uri? captureUrl) ? captureUrl : null;
        TraceEntries = Array.AsReadOnly(response?.TraceEntries ?? Array.Empty<string>());
        ExpectedRendererSha256 = rendererSha256;
        ExpectedWorkerSha256 = workerSha256;
        ExpectedRendererFilesSha256 = rendererFilesSha256;
        ExpectedWorkerFilesSha256 = workerFilesSha256;
        ReportedRendererSha256 = response?.RendererSha256;
        ReportedWorkerSha256 = response?.WorkerSha256;
        ReportedRendererFilesSha256 = response?.RendererFilesSha256;
        ReportedWorkerFilesSha256 = response?.WorkerFilesSha256;
    }

    /// <summary>Last workflow phase reached before failure.</summary>
    public HtmlIsolatedPublicPagePhase Phase { get; }
    /// <summary>Successfully acquired resource provenance available before failure.</summary>
    public IReadOnlyList<HtmlPublicResourceEvidence> Resources { get; }
    /// <summary>Discovered resources omitted by policy before failure.</summary>
    public IReadOnlyList<HtmlPublicSkippedResource> SkippedResources { get; }
    /// <summary>Full immutable OCI image ID requested for the run.</summary>
    public string ImageId { get; }
    /// <summary>Versioned isolation profile attempted for the run.</summary>
    public HtmlPublicPageIsolationProfile IsolationProfile => HtmlPublicPageIsolationProfile.NetworklessRootlessOciV1;
    /// <summary>Concrete controls required for the attempted isolation profile.</summary>
    public string IsolationPolicy => "rootless-podman;seccomp;cgroups-cpu-memory-pids;network-none;read-only;uid-65532;cap-drop-all;no-new-privileges;no-mounts";
    /// <summary>Container identity when startup reached container creation.</summary>
    public string? ContainerName { get; }
    /// <summary>Whether removal was verified; <see langword="null"/> when no lease was returned.</summary>
    public bool? ContainerRemoved { get; }
    /// <summary>Bounded cleanup failure text when removal could not be verified.</summary>
    public string? CleanupError { get; }
    /// <summary>Inner runtime provider identity when the worker reported one.</summary>
    public string? ProviderId { get; }
    /// <summary>Captured URL when capture completed before a later failure.</summary>
    public Uri? CaptureUrl { get; }
    /// <summary>Bounded runtime trace summaries available before failure.</summary>
    public IReadOnlyList<string> TraceEntries { get; }
    /// <summary>Expected renderer entry-assembly digest calculated by the host, when available.</summary>
    public string? ExpectedRendererSha256 { get; }
    /// <summary>Expected script-worker entry-assembly digest calculated by the host, when available.</summary>
    public string? ExpectedWorkerSha256 { get; }
    /// <summary>Expected complete renderer payload digest calculated by the host, when available.</summary>
    public string? ExpectedRendererFilesSha256 { get; }
    /// <summary>Expected complete script-worker payload digest calculated by the host, when available.</summary>
    public string? ExpectedWorkerFilesSha256 { get; }
    /// <summary>Renderer entry-assembly digest reported from isolation, when available.</summary>
    public string? ReportedRendererSha256 { get; }
    /// <summary>Script-worker entry-assembly digest reported from isolation, when available.</summary>
    public string? ReportedWorkerSha256 { get; }
    /// <summary>Complete renderer payload digest reported from isolation, when available.</summary>
    public string? ReportedRendererFilesSha256 { get; }
    /// <summary>Complete script-worker payload digest reported from isolation, when available.</summary>
    public string? ReportedWorkerFilesSha256 { get; }
}

/// <summary>An isolated public-page run failed after admission and carries bounded partial evidence.</summary>
public sealed class HtmlIsolatedPublicPageException : InvalidOperationException {
    internal HtmlIsolatedPublicPageException(string message, Exception innerException,
        HtmlIsolatedPublicPageFailureEvidence evidence) : base(message, innerException) => Evidence = evidence;

    /// <summary>Partial acquisition, worker identity, trace, and cleanup evidence.</summary>
    public HtmlIsolatedPublicPageFailureEvidence Evidence { get; }
}

/// <summary>An isolated public-page run was canceled or exceeded its operation deadline and carries bounded partial evidence.</summary>
public sealed class HtmlIsolatedPublicPageCanceledException : OperationCanceledException {
    internal HtmlIsolatedPublicPageCanceledException(string message, Exception innerException,
        CancellationToken cancellationToken, HtmlIsolatedPublicPageFailureEvidence evidence)
        : base(message, innerException, cancellationToken) => Evidence = evidence;

    /// <summary>Partial acquisition, worker identity, trace, and cleanup evidence.</summary>
    public HtmlIsolatedPublicPageFailureEvidence Evidence { get; }
}

/// <summary>Verified successful acquisition and whole-pipeline isolated page result.</summary>
public sealed class HtmlIsolatedPublicPageResult {
    internal HtmlIsolatedPublicPageResult(string scenarioId, string sourceLicense, Uri requestedUrl,
        IReadOnlyList<HtmlPublicResourceEvidence> resources, IReadOnlyList<HtmlPublicSkippedResource> skippedResources,
        string imageId, string containerName, HtmlPublicRenderResponse response, byte[] screen, byte[] print, byte[] screenToPage) {
        ScenarioId = scenarioId;
        SourceLicense = sourceLicense;
        RequestedUrl = requestedUrl;
        Resources = resources;
        SkippedResources = skippedResources;
        ImageId = imageId;
        ContainerName = containerName;
        ProviderId = response.ProviderId ?? string.Empty;
        CaptureUrl = new Uri(response.CaptureUrl ?? throw new HtmlScriptRuntimeException("The isolated capture URL is missing."));
        CaptureManifest = response.CaptureManifest ?? string.Empty;
        TraceEntries = Array.AsReadOnly(response.TraceEntries ?? Array.Empty<string>());
        RendererSha256 = response.RendererSha256;
        WorkerSha256 = response.WorkerSha256;
        RendererFilesSha256 = response.RendererFilesSha256;
        WorkerFilesSha256 = response.WorkerFilesSha256;
        Outputs = Array.AsReadOnly(new[] {
            new HtmlIsolatedPageOutput("screen.png", "image/png", screen),
            new HtmlIsolatedPageOutput("print.pdf", "application/pdf", print),
            new HtmlIsolatedPageOutput("screen-to-page.pdf", "application/pdf", screenToPage)
        });
        UnsupportedFeatures = Array.AsReadOnly(new[] {
            "cross-origin-frame-execution",
            "child-frame-module-graphs",
            "module-import-attributes",
            "xhr-custom-request-headers",
            "dynamic-non-get-requests",
            "cookies-and-credentials"
        });
    }

    /// <summary>Stable caller-provided scenario identity.</summary>
    public string ScenarioId { get; }
    /// <summary>Caller-provided source license or usage-rights note.</summary>
    public string SourceLicense { get; }
    /// <summary>Originally requested page URL.</summary>
    public Uri RequestedUrl { get; }
    /// <summary>Acquired document and resource provenance.</summary>
    public IReadOnlyList<HtmlPublicResourceEvidence> Resources { get; }
    /// <summary>Discovered resources omitted by the acquisition policy.</summary>
    public IReadOnlyList<HtmlPublicSkippedResource> SkippedResources { get; }
    /// <summary>Full immutable OCI image ID.</summary>
    public string ImageId { get; }
    /// <summary>Unique container identity used for the run.</summary>
    public string ContainerName { get; }
    /// <summary>Confirms that successful return happened after verified container removal.</summary>
    public bool ContainerRemoved => true;
    /// <summary>Versioned isolation profile applied to the whole parsing, scripting, capture and rendering pipeline.</summary>
    public HtmlPublicPageIsolationProfile IsolationProfile => HtmlPublicPageIsolationProfile.NetworklessRootlessOciV1;
    /// <summary>Concrete controls enforced and inspected for this profile.</summary>
    public string IsolationPolicy => "rootless-podman;seccomp;cgroups-cpu-memory-pids;network-none;read-only;uid-65532;cap-drop-all;no-new-privileges;no-mounts";
    /// <summary>Runtime provider identity reported from inside isolation.</summary>
    public string ProviderId { get; }
    /// <summary>Final captured document URL.</summary>
    public Uri CaptureUrl { get; }
    /// <summary>Content-addressed capture manifest identity.</summary>
    public string CaptureManifest { get; }
    /// <summary>Bounded runtime trace summaries retained by the isolated renderer.</summary>
    public IReadOnlyList<string> TraceEntries { get; }
    /// <summary>Known browser features outside this versioned profile.</summary>
    public IReadOnlyList<string> UnsupportedFeatures { get; }
    /// <summary>SHA-256 digest of the renderer entry assembly returned from isolation.</summary>
    public string RendererSha256 { get; }
    /// <summary>SHA-256 digest of the script-worker entry assembly returned from isolation.</summary>
    public string WorkerSha256 { get; }
    /// <summary>Deterministic digest of the complete published renderer payload.</summary>
    public string RendererFilesSha256 { get; }
    /// <summary>Deterministic digest of the complete published script-worker payload.</summary>
    public string WorkerFilesSha256 { get; }
    /// <summary>Screen, browser-print and screen-to-page artifacts.</summary>
    public IReadOnlyList<HtmlIsolatedPageOutput> Outputs { get; }
}
