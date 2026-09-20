using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Html.Runtime;

namespace OfficeIMO.Html.Runtime.Rendering;

/// <summary>Acquires approved public resources on the host and performs all parsing, scripting, capture, and rendering in a verified networkless OCI worker.</summary>
public static class HtmlIsolatedPublicPageWorkflow {
    /// <summary>Runs one named public page through screen, print, and screen-to-page output.</summary>
    public static async Task<HtmlIsolatedPublicPageResult> RunAsync(
        HtmlIsolatedPublicPageExecutionOptions execution,
        HtmlIsolatedPublicPageRequest request,
        CancellationToken cancellationToken = default) => await RunCoreAsync(execution, request, null, cancellationToken).ConfigureAwait(false);

    internal static Task<HtmlIsolatedPublicPageResult> RunWithTransportAsync(
        HtmlIsolatedPublicPageExecutionOptions execution,
        HtmlIsolatedPublicPageRequest request,
        HtmlPublicAcquisitionTransport transport,
        CancellationToken cancellationToken = default) =>
        RunCoreAsync(execution, request, transport ?? throw new ArgumentNullException(nameof(transport)), cancellationToken);

    private static async Task<HtmlIsolatedPublicPageResult> RunCoreAsync(
        HtmlIsolatedPublicPageExecutionOptions execution,
        HtmlIsolatedPublicPageRequest request,
        HtmlPublicAcquisitionTransport? transport,
        CancellationToken cancellationToken) {
        HtmlIsolatedPublicPageExecutionOptions.Snapshot run = (execution ?? throw new ArgumentNullException(nameof(execution))).Validate();
        HtmlIsolatedPublicPageRequest.Snapshot input = (request ?? throw new ArgumentNullException(nameof(request))).Validate();
        using var deadline = new CancellationTokenSource(run.OperationTimeout);
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, deadline.Token);
        var assets = new List<HtmlPublicResourceResult>();
        var skipped = new List<HtmlPublicSkippedResource>();
        HtmlPublicResourceResult? document = null;
        HtmlPublicRenderResponse? response = null;
        byte[]? screen = null, print = null, screenToPage = null;
        HtmlOciWorkerLease? lease = null;
        string? containerName = null;
        bool? containerRemoved = null;
        string? cleanupError = null;
        string? expectedRenderer = null, expectedWorker = null;
        string? expectedRendererFiles = null, expectedWorkerFiles = null;
        HtmlPublicFontPackageIdentity? expectedFontPackage = null;
        Exception? failure = null;
        HtmlIsolatedPublicPagePhase phase = HtmlIsolatedPublicPagePhase.Admission;

        try {
            HtmlRuntimeResourcePolicy callerPolicy = input.Runtime.ResourcePolicy;
            int acquisitionRequests = Math.Min(32, callerPolicy.MaxRequests);
            int assetLimit = Math.Min(24, Math.Max(0, acquisitionRequests - 1));
            if (input.SeedResourceUrls.Length > assetLimit)
                throw new ArgumentException("Seed resources exceed the caller's acquisition request limit.", nameof(request));
            long acquisitionResourceBytes = Math.Min(4L * 1024 * 1024,
                Math.Min(callerPolicy.MaxResourceBytes, callerPolicy.MaxTotalBytes));
            long acquisitionTotalBytes = Math.Min(16L * 1024 * 1024, callerPolicy.MaxTotalBytes);
            string[] allowedHosts = new[] { input.Url.IdnHost }
                .Concat(input.SeedResourceUrls.Select(resource => resource.IdnHost))
                .Concat(input.AllowedHosts).ToArray();
            TimeSpan acquisitionTimeout = callerPolicy.Timeout < TimeSpan.FromSeconds(20)
                ? callerPolicy.Timeout : TimeSpan.FromSeconds(20);
            var broker = transport == null
                ? new HtmlPublicResourceBroker(allowedHosts, maxRequests: acquisitionRequests,
                    maxResourceBytes: acquisitionResourceBytes, maxTotalBytes: acquisitionTotalBytes,
                    timeout: acquisitionTimeout, maxRedirects: Math.Min(callerPolicy.MaxRedirects, 5),
                    maxRequestBytes: Math.Min(callerPolicy.MaxRequestBytes, 1024L * 1024),
                    maxTotalRequestBytes: Math.Min(callerPolicy.MaxTotalRequestBytes, 16L * 1024 * 1024),
                    dynamicHosts: input.AllowedDynamicRequestOrigins.Select(origin => origin.IdnHost))
                : new HtmlPublicResourceBroker(allowedHosts, transport.ResolveAddresses, transport.Connect,
                    maxRequests: acquisitionRequests, maxResourceBytes: acquisitionResourceBytes,
                    maxTotalBytes: acquisitionTotalBytes, timeout: acquisitionTimeout,
                    maxRedirects: Math.Min(callerPolicy.MaxRedirects, 5),
                    maxRequestBytes: Math.Min(callerPolicy.MaxRequestBytes, 1024L * 1024),
                    maxTotalRequestBytes: Math.Min(callerPolicy.MaxTotalRequestBytes, 16L * 1024 * 1024),
                    dynamicHosts: input.AllowedDynamicRequestOrigins.Select(origin => origin.IdnHost));

            phase = HtmlIsolatedPublicPagePhase.Acquisition;
            document = await broker.FetchAsync(input.Url, operation.Token).ConfigureAwait(false);
            string html = HtmlPublicResourceBroker.DecodeUtf8Html(document.Resource);
            foreach (Uri resourceUrl in input.SeedResourceUrls.DistinctBy(
                         HtmlRuntimeResourcePolicy.Key, StringComparer.Ordinal)) {
                assets.Add(await broker.FetchAsync(resourceUrl, operation.Token).ConfigureAwait(false));
            }

            phase = HtmlIsolatedPublicPagePhase.Preflight;
            expectedRenderer = Digest(await File.ReadAllBytesAsync(run.RendererPath, operation.Token).ConfigureAwait(false));
            expectedWorker = Digest(await File.ReadAllBytesAsync(run.WorkerPath, operation.Token).ConfigureAwait(false));
            expectedRendererFiles = HtmlPublicArtifactDigest.DirectorySha256(Path.GetDirectoryName(run.RendererPath)!);
            expectedWorkerFiles = HtmlPublicArtifactDigest.DirectorySha256(Path.GetDirectoryName(run.WorkerPath)!);
            expectedFontPackage = HtmlPublicFontPackage.Load(Path.GetDirectoryName(run.RendererPath)!);

            HtmlScriptRequest page = input.Runtime;
            page.Html = html;
            page.DocumentUrl = document.Resource.FinalUrl;
            page.Resources = assets.Select(asset => asset.Resource).ToArray();
            page.ResourcePolicy = BoundedPolicy(callerPolicy, broker.AllowedOrigins
                .Where(origin => !origin.IdnHost.Equals(page.DocumentUrl.IdnHost, StringComparison.OrdinalIgnoreCase))
                .ToArray());
            page.FailOnFetchReplayDiscovery = true;
            page = page.Snapshot();
            var renderRequest = new HtmlPublicRenderRequest { Page = page, Actions = input.Actions,
                FinalReadyExpression = input.FinalReadyExpression,
                MaxOutputBytesPerArtifact = input.MaxOutputBytesPerArtifact,
                MaxTotalOutputBytes = input.MaxTotalOutputBytes };
            if (HtmlRuntimeProtocol.MeasureCharacters(renderRequest) > HtmlPublicRenderRequest.MaximumProtocolCharacters)
                throw new ArgumentException("The acquired page and automation request exceed the isolated protocol budget.", nameof(request));

            phase = HtmlIsolatedPublicPagePhase.IsolatedStartup;
            lease = await HtmlOciWorkerLease.StartAsync(
                run.Executable, run.Arguments, run.ImageId, operation.Token).ConfigureAwait(false);
            containerName = lease.ContainerName;
            containerRemoved = false;
            Process process = lease.Process;
            Task<string> stderr = DrainErrorAsync(process.StandardError.BaseStream, operation.Token);
            await HtmlRuntimeProtocol.WriteAsync(process.StandardInput.BaseStream, renderRequest,
                HtmlPublicRenderRequest.MaximumProtocolCharacters, operation.Token).ConfigureAwait(false);
            var known = new HashSet<string>(StringComparer.Ordinal) {
                HtmlRuntimeResourcePolicy.Key(document.Resource.Url),
                HtmlRuntimeResourcePolicy.Key(document.Resource.FinalUrl)
            };
            foreach (HtmlPublicResourceResult asset in assets) {
                known.Add(HtmlRuntimeResourcePolicy.Key(asset.Resource.Url));
                known.Add(HtmlRuntimeResourcePolicy.Key(asset.Resource.FinalUrl));
            }
            var knownDynamic = new HashSet<string>(StringComparer.Ordinal);

            for (int round = 0; ; round++) {
                phase = HtmlIsolatedPublicPagePhase.ResourceDiscovery;
                response = await HtmlRuntimeProtocol.ReadAsync<HtmlPublicRenderResponse>(
                    process.StandardOutput.BaseStream, 24 * 1024 * 1024, operation.Token).ConfigureAwait(false)
                    ?? throw new HtmlScriptRuntimeException("The isolated renderer returned no discovery response.");
                phase = WorkerPhase(response.Stage);
                VerifyIdentity(response, expectedRenderer, expectedWorker, expectedRendererFiles, expectedWorkerFiles,
                    expectedFontPackage);
                ThrowWorkerError(response);
                if (response.DiscoveryComplete) break;
                if (round >= 16 || response.DiscoveryUrls == null || response.DiscoveryRequests == null ||
                    (long)response.DiscoveryUrls.Length + response.DiscoveryRequests.Length is 0 or > 128)
                    throw new HtmlScriptRuntimeException("The isolated renderer exceeded the resource discovery limit.");
                var batch = new List<HtmlRuntimeResource>();
                var fetchBatch = new List<HtmlRuntimeFetchReplay>();
                foreach (string candidate in response.DiscoveryUrls) {
                    if (!Uri.TryCreate(candidate, UriKind.Absolute, out Uri? resourceUrl)) {
                        skipped.Add(new HtmlPublicSkippedResource(candidate, "invalid-url"));
                        continue;
                    }
                    try { HtmlPublicResourceBroker.ValidateUrl(resourceUrl); }
                    catch (ArgumentException) {
                        skipped.Add(new HtmlPublicSkippedResource(candidate, "unsupported-url"));
                        continue;
                    }
                    if (!broker.AllowsHost(resourceUrl)) {
                        skipped.Add(new HtmlPublicSkippedResource(candidate, "host-not-allowed"));
                        continue;
                    }
                    if (!known.Add(HtmlRuntimeResourcePolicy.Key(resourceUrl))) continue;
                    if (assets.Count >= assetLimit) {
                        skipped.Add(new HtmlPublicSkippedResource(candidate, "resource-count-limit"));
                        continue;
                    }
                    HtmlPublicResourceResult asset = await broker.FetchAsync(resourceUrl, operation.Token).ConfigureAwait(false);
                    assets.Add(asset);
                    known.Add(HtmlRuntimeResourcePolicy.Key(asset.Resource.FinalUrl));
                    batch.Add(asset.Resource);
                }
                foreach (HtmlRuntimeFetchDiscovery discovery in response.DiscoveryRequests) {
                    HtmlRuntimeFetchRequest dynamic = discovery?.Request
                        ?? throw new HtmlScriptRuntimeException("The isolated renderer returned an invalid dynamic request.");
                    if (!knownDynamic.Add(discovery.Identity))
                        throw new HtmlScriptRuntimeException("The isolated renderer repeated a supplied dynamic request occurrence.");
                    if (!input.AllowedDynamicRequestMethods.Contains(dynamic.Method, StringComparer.Ordinal))
                        throw new HtmlScriptRuntimeException("Dynamic request method " + dynamic.Method + " was not authorized by the caller.");
                    string dynamicOrigin = HtmlRuntimeResourcePolicy.Origin(dynamic.Url);
                    if (dynamicOrigin != HtmlRuntimeResourcePolicy.Origin(page.DocumentUrl) &&
                        !input.AllowedDynamicRequestOrigins.Any(origin => HtmlRuntimeResourcePolicy.Origin(origin) == dynamicOrigin))
                        throw new HtmlScriptRuntimeException("The dynamic request origin was not authorized by the caller.");
                    HtmlPublicResourceBroker.ValidateDynamicRequest(dynamic);
                    if (assets.Count >= assetLimit)
                        throw new HtmlScriptRuntimeException("The dynamic acquisition resource count limit was exceeded.");
                    HtmlPublicResourceResult asset = await broker.FetchAsync(discovery, page.DocumentUrl, operation.Token).ConfigureAwait(false);
                    assets.Add(asset);
                    fetchBatch.Add(new HtmlRuntimeFetchReplay(dynamic, discovery.Occurrence,
                        asset.DynamicHops ?? throw new HtmlScriptRuntimeException("The dynamic acquisition transcript is missing.")));
                }
                await HtmlRuntimeProtocol.WriteAsync(process.StandardInput.BaseStream,
                    new HtmlPublicResourceBatch { Resources = batch.ToArray(), FetchReplays = fetchBatch.ToArray() },
                    24 * 1024 * 1024, operation.Token).ConfigureAwait(false);
            }

            phase = HtmlIsolatedPublicPagePhase.Rendering;
            process.StandardInput.Close();
            response = await HtmlRuntimeProtocol.ReadAsync<HtmlPublicRenderResponse>(
                process.StandardOutput.BaseStream, 24 * 1024 * 1024, operation.Token).ConfigureAwait(false)
                ?? throw new HtmlScriptRuntimeException("The isolated renderer returned no final response.");
            phase = WorkerPhase(response.Stage);
            await process.WaitForExitAsync(operation.Token).ConfigureAwait(false);
            string errorOutput = await stderr.ConfigureAwait(false);
            if (process.ExitCode != 0)
                throw new HtmlScriptRuntimeException("The isolated renderer exited " + process.ExitCode + ": " + errorOutput);
            VerifyIdentity(response, expectedRenderer, expectedWorker, expectedRendererFiles, expectedWorkerFiles,
                expectedFontPackage);
            ThrowWorkerError(response);
            phase = HtmlIsolatedPublicPagePhase.Output;
            screen = response.Screen ?? throw new HtmlScriptRuntimeException("The isolated screen output is missing.");
            print = response.Print ?? throw new HtmlScriptRuntimeException("The isolated print output is missing.");
            screenToPage = response.ScreenToPage ?? throw new HtmlScriptRuntimeException("The isolated screen-to-page output is missing.");
            if (screen.LongLength > input.MaxOutputBytesPerArtifact || print.LongLength > input.MaxOutputBytesPerArtifact ||
                screenToPage.LongLength > input.MaxOutputBytesPerArtifact ||
                (long)screen.Length + print.Length + screenToPage.Length > input.MaxTotalOutputBytes)
                throw new HtmlScriptRuntimeException("The isolated render output exceeds its byte budget.");
        } catch (HtmlOciWorkerStartException error) {
            containerName = error.ContainerName;
            containerRemoved = error.ContainerRemoved;
            cleanupError = error.CleanupError;
            failure = error.InnerException ?? error;
        } catch (Exception error) {
            failure = error;
        } finally {
            if (lease != null) {
                try {
                    await lease.DisposeAsync().ConfigureAwait(false);
                    containerRemoved = true;
                } catch (Exception error) {
                    containerRemoved = false;
                    cleanupError = BoundedMessage(error);
                    if (failure == null) {
                        phase = HtmlIsolatedPublicPagePhase.Cleanup;
                        failure = error;
                    }
                }
            }
        }

        if (failure != null) {
            HtmlIsolatedPublicPageFailureEvidence evidence = FailureEvidence(phase, input, document, assets, skipped,
                run.ImageId, containerName, containerRemoved, cleanupError, response, expectedRenderer, expectedWorker,
                expectedRendererFiles, expectedWorkerFiles, expectedFontPackage);
            if (failure is OperationCanceledException && phase != HtmlIsolatedPublicPagePhase.Cleanup) {
                string message = cancellationToken.IsCancellationRequested
                    ? "The isolated public-page run was canceled."
                    : "The isolated public-page operation deadline was exceeded.";
                throw new HtmlIsolatedPublicPageCanceledException(message, failure,
                    cancellationToken.IsCancellationRequested ? cancellationToken : operation.Token, evidence);
            }
            throw new HtmlIsolatedPublicPageException("The isolated public-page run failed during " +
                phase.ToString().ToLowerInvariant() + ".", failure, evidence);
        }

        var evidenceResources = new[] { document! }.Concat(assets)
            .Select(resource => new HtmlPublicResourceEvidence(resource, input.RetainInputBytes)).ToArray();
        return new HtmlIsolatedPublicPageResult(input.ScenarioId, input.SourceLicense, input.Url,
            Array.AsReadOnly(evidenceResources), Array.AsReadOnly(skipped.ToArray()), run.ImageId,
            containerName ?? throw new HtmlScriptRuntimeException("The isolated container identity is missing."),
            response!, screen!, print!, screenToPage!);
    }

    internal static HtmlRuntimeResourcePolicy BoundedPolicy(HtmlRuntimeResourcePolicy caller, Uri[] allowedOrigins) => new() {
        AllowNetwork = false,
        AllowedOrigins = allowedOrigins,
        Timeout = caller.Timeout < TimeSpan.FromSeconds(5) ? caller.Timeout : TimeSpan.FromSeconds(5),
        MaxConcurrentRequests = Math.Min(caller.MaxConcurrentRequests, 4),
        MaxRequests = Math.Min(caller.MaxRequests, 64),
        MaxResourceBytes = Math.Min(caller.MaxResourceBytes, 4L * 1024 * 1024),
        MaxTotalBytes = Math.Min(caller.MaxTotalBytes, 16L * 1024 * 1024),
        MaxRequestBytes = Math.Min(caller.MaxRequestBytes, 1024L * 1024),
        MaxTotalRequestBytes = Math.Min(caller.MaxTotalRequestBytes, 16L * 1024 * 1024),
        MaxRedirects = Math.Min(caller.MaxRedirects, 5)
    };

    private static HtmlIsolatedPublicPageFailureEvidence FailureEvidence(HtmlIsolatedPublicPagePhase phase,
        HtmlIsolatedPublicPageRequest.Snapshot input, HtmlPublicResourceResult? document,
        IEnumerable<HtmlPublicResourceResult> assets, IEnumerable<HtmlPublicSkippedResource> skipped,
        string imageId, string? containerName, bool? containerRemoved, string? cleanupError,
        HtmlPublicRenderResponse? response, string? renderer, string? worker, string? rendererFiles, string? workerFiles,
        HtmlPublicFontPackageIdentity? expectedFontPackage) {
        IEnumerable<HtmlPublicResourceResult> acquired = document == null
            ? assets
            : new[] { document }.Concat(assets);
        var resources = acquired.Select(resource => new HtmlPublicResourceEvidence(resource, input.RetainInputBytes)).ToArray();
        return new HtmlIsolatedPublicPageFailureEvidence(phase, Array.AsReadOnly(resources),
            Array.AsReadOnly(skipped.ToArray()), imageId, containerName, containerRemoved, cleanupError,
            response, renderer, worker, expectedFontPackage, rendererFiles, workerFiles);
    }

    private static void ThrowWorkerError(HtmlPublicRenderResponse response) {
        if (response.Error != null) throw new HtmlScriptRuntimeException(response.ErrorKind + ": " + response.Error);
    }

    internal static HtmlIsolatedPublicPagePhase WorkerPhase(HtmlPublicRenderStage stage) => stage switch {
        HtmlPublicRenderStage.ResourceDiscovery => HtmlIsolatedPublicPagePhase.ResourceDiscovery,
        HtmlPublicRenderStage.Rendering => HtmlIsolatedPublicPagePhase.Rendering,
        HtmlPublicRenderStage.Output => HtmlIsolatedPublicPagePhase.Output,
        _ => throw new HtmlScriptRuntimeException("The isolated renderer returned an invalid workflow stage.")
    };

    private static void VerifyIdentity(HtmlPublicRenderResponse response, string renderer, string worker,
        string rendererFiles, string workerFiles, HtmlPublicFontPackageIdentity expectedFontPackage) {
        if (!response.RendererSha256.Equals(renderer, StringComparison.OrdinalIgnoreCase)
            || !response.WorkerSha256.Equals(worker, StringComparison.OrdinalIgnoreCase)
            || !response.RendererFilesSha256.Equals(rendererFiles, StringComparison.OrdinalIgnoreCase)
            || !response.WorkerFilesSha256.Equals(workerFiles, StringComparison.OrdinalIgnoreCase)
            || !HtmlPublicFontPackage.Matches(expectedFontPackage, response.FontPackage)) {
            throw new HtmlScriptRuntimeException(
                "The isolated image does not contain the expected renderer, script worker, and font-package payloads.");
        }
    }

    private static async Task<string> DrainErrorAsync(Stream stream, CancellationToken token) {
        var prefix = new MemoryStream();
        var buffer = new byte[4096];
        int read;
        while ((read = await stream.ReadAsync(buffer, token).ConfigureAwait(false)) != 0) {
            if (prefix.Length < 2048) prefix.Write(buffer, 0, (int)Math.Min(read, 2048 - prefix.Length));
        }
        return Encoding.UTF8.GetString(prefix.ToArray());
    }

    private static string Digest(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static string BoundedMessage(Exception error) => error.Message.Length <= 1024 ? error.Message : error.Message[..1024];
}
