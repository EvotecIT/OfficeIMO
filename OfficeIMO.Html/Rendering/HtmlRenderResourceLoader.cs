using System.Security.Cryptography;

namespace OfficeIMO.Html;

/// <summary>
/// Operation-scoped owner for HTML resource policy, resolution, MIME validation, deduplication,
/// caching, budgets, timeouts, cancellation evidence, canonical identities, and content digests.
/// </summary>
public sealed class HtmlResourceSession {
    private readonly Dictionary<string, HtmlResolvedResource> _resources = new Dictionary<string, HtmlResolvedResource>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _resolvedSources = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _attempted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _budgetedStylesheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _rejectedStylesheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private readonly List<HtmlResourceSessionEntry> _entries = new List<HtmlResourceSessionEntry>();
    private readonly IReadOnlyList<HtmlResourceSessionEntry> _readOnlyEntries;

    /// <summary>Creates an independent resource session.</summary>
    public HtmlResourceSession(
        HtmlUrlPolicy? resourcePolicy = null,
        HtmlRenderResourceResolver? resolver = null,
        TimeSpan? resourceTimeout = null,
        int maxConcurrentLoads = 8,
        long maxResourceBytes = 10L * 1024L * 1024L,
        long maxTotalResourceBytes = 50L * 1024L * 1024L,
        int maxResourceCount = 256,
        int maxResourceRequests = 512,
        int maxStylesheetImportDepth = 16) {
        ResourcePolicy = (resourcePolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone();
        Resolver = resolver;
        ResourceTimeout = resourceTimeout ?? TimeSpan.FromSeconds(30D);
        MaxConcurrentLoads = maxConcurrentLoads;
        MaxResourceBytes = maxResourceBytes;
        MaxTotalResourceBytes = maxTotalResourceBytes;
        MaxResourceCount = maxResourceCount;
        MaxResourceRequests = maxResourceRequests;
        MaxStylesheetImportDepth = maxStylesheetImportDepth;
        ValidateLimits();
        Diagnostics = new HtmlDiagnosticReport();
        _readOnlyEntries = _entries.AsReadOnly();
    }

    internal HtmlResourceSession(HtmlRenderOptions options, HtmlDiagnosticReport diagnostics) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        ResourcePolicy = options.GetResourceUrlPolicy().Clone();
        Resolver = options.ResourceResolver;
        SynchronousResolver = options.SynchronousResourceResolver;
        ResourceTimeout = options.ResourceTimeout;
        MaxConcurrentLoads = options.MaxConcurrentResourceLoads;
        MaxResourceBytes = options.MaxResourceBytes;
        MaxTotalResourceBytes = options.MaxTotalResourceBytes;
        MaxResourceCount = options.MaxResourceCount;
        MaxResourceRequests = options.MaxResourceRequests;
        MaxStylesheetImportDepth = options.MaxStylesheetImportDepth;
        ValidateLimits();
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        _readOnlyEntries = _entries.AsReadOnly();
    }

    /// <summary>Resource URL policy snapshot owned by this operation.</summary>
    public HtmlUrlPolicy ResourcePolicy { get; }

    /// <summary>Asynchronous application resolver snapshot owned by this operation.</summary>
    public HtmlRenderResourceResolver? Resolver { get; }

    /// <summary>Maximum duration of one resolver invocation.</summary>
    public TimeSpan ResourceTimeout { get; }

    /// <summary>Maximum concurrent asynchronous resolver invocations.</summary>
    public int MaxConcurrentLoads { get; }

    /// <summary>Maximum accepted bytes for one resource.</summary>
    public long MaxResourceBytes { get; }

    /// <summary>Maximum accepted bytes for the operation.</summary>
    public long MaxTotalResourceBytes { get; }

    /// <summary>Maximum accepted resource count.</summary>
    public int MaxResourceCount { get; }

    /// <summary>Maximum attempted resolver requests.</summary>
    public int MaxResourceRequests { get; }

    /// <summary>Maximum external stylesheet import depth.</summary>
    public int MaxStylesheetImportDepth { get; }

    /// <summary>Diagnostics emitted by this session.</summary>
    public HtmlDiagnosticReport Diagnostics { get; }

    /// <summary>Accepted canonical resources in acceptance order.</summary>
    public IReadOnlyList<HtmlResourceSessionEntry> Resources => _readOnlyEntries;

    internal HtmlRenderSynchronousResourceResolver? SynchronousResolver { get; }

    /// <summary>Total accepted encoded bytes.</summary>
    public long AcceptedResourceBytes { get; private set; }

    /// <summary>Number of accepted deduplicated resources.</summary>
    public int AcceptedResourceCount { get; private set; }

    /// <summary>Number of resolver requests reserved by the session.</summary>
    public int ResolverRequestCount { get; private set; }

    /// <summary>Resolves a manifest synchronously using the configured synchronous package resolver, when present.</summary>
    public static HtmlResourceSession Resolve(
        HtmlResourceManifest manifest,
        HtmlRenderOptions options,
        HtmlDiagnosticReport? diagnostics = null,
        CancellationToken cancellationToken = default) =>
        HtmlRenderResourceLoader.Load(manifest, options, diagnostics ?? new HtmlDiagnosticReport(), cancellationToken);

    /// <summary>Resolves a manifest asynchronously using one operation-scoped session.</summary>
    public static Task<HtmlResourceSession> ResolveAsync(
        HtmlResourceManifest manifest,
        HtmlRenderOptions options,
        HtmlDiagnosticReport? diagnostics = null,
        CancellationToken cancellationToken = default) =>
        HtmlRenderResourceLoader.LoadAsync(manifest, options, diagnostics ?? new HtmlDiagnosticReport(), cancellationToken);

    internal void MarkAttempted(HtmlResourceReference reference) {
        if (reference.Source.Length > 0) _attempted.Add(reference.Source);
        if (reference.ResolvedSource.Length > 0) _attempted.Add(reference.ResolvedSource);
    }

    internal void Add(HtmlResourceReference reference, HtmlResolvedResource resource) {
        AcceptedResourceBytes += resource.Length;
        AcceptedResourceCount++;
        if (reference.Source.Length > 0) {
            _resources[reference.Source] = resource;
            _resolvedSources[reference.Source] = reference.ResolvedSource;
        }

        if (reference.ResolvedSource.Length > 0) {
            _resources[reference.ResolvedSource] = resource;
            _resolvedSources[reference.ResolvedSource] = reference.ResolvedSource;
        }
        _entries.Add(CreateEntry(reference.Kind, reference.Source, reference.ResolvedSource, resource));
    }

    internal bool TryReserveRequest(HtmlResourceReference reference) {
        if (ResolverRequestCount < MaxResourceRequests) {
            ResolverRequestCount++;
            return true;
        }

        Diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.ResourceRequestLimitExceeded,
            "Resource resolver invocations exceeded the configured operation-wide request limit.",
            HtmlDiagnosticSeverity.Error, reference.Source, "limit=" + MaxResourceRequests,
            HtmlConversionLossKind.Omission);
        return false;
    }

    internal bool TryAccept(HtmlResourceReference reference, HtmlResolvedResource resource, out bool stop) {
        stop = false;
        long length = resource.Length;
        if (length > MaxResourceBytes) {
            Diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded,
                "A resolved resource exceeded the configured per-resource byte limit.",
                HtmlDiagnosticSeverity.Warning, reference.Source, "bytes=" + length,
                HtmlConversionLossKind.Omission);
            return false;
        }
        if (AcceptedResourceCount >= MaxResourceCount) {
            Diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded,
                "Resolved resources exceeded the configured operation-wide count limit.",
                HtmlDiagnosticSeverity.Error, reference.Source, "limit=" + MaxResourceCount,
                HtmlConversionLossKind.Omission);
            stop = true;
            return false;
        }
        if (length > MaxTotalResourceBytes - AcceptedResourceBytes) {
            Diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded,
                "Resolved resources exceeded the configured total byte limit.",
                HtmlDiagnosticSeverity.Error, reference.Source,
                "bytes=" + (AcceptedResourceBytes + length), HtmlConversionLossKind.Omission);
            stop = true;
            return false;
        }
        if (!IsAcceptedContentType(reference.Kind, resource.ContentType)) {
            Diagnostics.Add("OfficeIMO.Html.Renderer", HtmlRenderDiagnosticCodes.ResourceContentTypeRejected,
                "A resolver returned an incompatible media type for the requested resource kind.",
                HtmlDiagnosticSeverity.Warning, reference.Source,
                reference.Kind + ":" + resource.ContentType, HtmlConversionLossKind.Omission);
            return false;
        }

        Add(reference, resource);
        return true;
    }

    internal bool CanAcceptInlineResource(
        long estimatedBytes,
        out string diagnosticCode,
        out string diagnosticDetail) {
        if (estimatedBytes > MaxResourceBytes) {
            diagnosticCode = HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded;
            diagnosticDetail = "bytes=" + estimatedBytes;
            return false;
        }

        if (AcceptedResourceCount >= MaxResourceCount) {
            diagnosticCode = HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded;
            diagnosticDetail = "limit=" + MaxResourceCount;
            return false;
        }

        if (estimatedBytes > MaxTotalResourceBytes - AcceptedResourceBytes) {
            diagnosticCode = HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded;
            diagnosticDetail = "bytes=" + (AcceptedResourceBytes + estimatedBytes);
            return false;
        }

        diagnosticCode = string.Empty;
        diagnosticDetail = string.Empty;
        return true;
    }

    internal bool TryAcceptInline(
        HtmlResourceKind kind,
        string resolvedSource,
        HtmlResolvedResource resource,
        out string diagnosticCode,
        out string diagnosticDetail) {
        if (string.IsNullOrWhiteSpace(resolvedSource)) throw new ArgumentException("A canonical inline resource source is required.", nameof(resolvedSource));
        if (resource == null) throw new ArgumentNullException(nameof(resource));
        if (_resources.ContainsKey(resolvedSource)) {
            diagnosticCode = string.Empty;
            diagnosticDetail = string.Empty;
            return true;
        }
        if (!IsAcceptedContentType(kind, resource.ContentType)) {
            diagnosticCode = HtmlRenderDiagnosticCodes.ResourceContentTypeRejected;
            diagnosticDetail = kind + ":" + resource.ContentType;
            return false;
        }
        if (!CanAcceptInlineResource(resource.Length, out diagnosticCode, out diagnosticDetail)) return false;
        AcceptedResourceBytes += resource.Length;
        AcceptedResourceCount++;
        _resources[resolvedSource] = resource;
        _resolvedSources[resolvedSource] = resolvedSource;
        _entries.Add(CreateEntry(kind, resolvedSource, resolvedSource, resource));
        diagnosticCode = string.Empty;
        diagnosticDetail = string.Empty;
        return true;
    }

    /// <summary>Gets a cached resource by original or canonical source.</summary>
    public bool TryGet(string? source, string? resolvedSource, out HtmlResolvedResource resource) {
        if (!string.IsNullOrWhiteSpace(resolvedSource) && _resources.TryGetValue(resolvedSource!, out resource!)) return true;
        if (!string.IsNullOrWhiteSpace(source) && _resources.TryGetValue(source!, out resource!)) return true;
        resource = null!;
        return false;
    }

    internal bool TryGetResolvedSource(string? source, string? resolvedSource, out string value) {
        if (!string.IsNullOrWhiteSpace(resolvedSource) && _resolvedSources.TryGetValue(resolvedSource!, out value!)) return true;
        if (!string.IsNullOrWhiteSpace(source) && _resolvedSources.TryGetValue(source!, out value!)) return true;
        value = string.Empty;
        return false;
    }

    internal bool WasAttempted(string? source, string? resolvedSource) =>
        !string.IsNullOrWhiteSpace(source) && _attempted.Contains(source!)
        || !string.IsNullOrWhiteSpace(resolvedSource) && _attempted.Contains(resolvedSource!);

    internal void MarkStylesheetBudgeted(HtmlResourceReference reference) =>
        MarkStylesheetState(_budgetedStylesheets, reference);

    internal void MarkStylesheetRejected(HtmlResourceReference reference) =>
        MarkStylesheetState(_rejectedStylesheets, reference);

    internal bool WasStylesheetBudgeted(string? source, string? resolvedSource) =>
        ContainsStylesheetState(_budgetedStylesheets, source, resolvedSource);

    internal bool WasStylesheetRejected(string? source, string? resolvedSource) =>
        ContainsStylesheetState(_rejectedStylesheets, source, resolvedSource);

    private static void MarkStylesheetState(HashSet<string> state, HtmlResourceReference reference) {
        if (reference.Source.Length > 0) state.Add(reference.Source);
        if (reference.ResolvedSource.Length > 0) state.Add(reference.ResolvedSource);
    }

    private static bool ContainsStylesheetState(HashSet<string> state, string? source, string? resolvedSource) =>
        !string.IsNullOrWhiteSpace(source) && state.Contains(source!)
        || !string.IsNullOrWhiteSpace(resolvedSource) && state.Contains(resolvedSource!);

    private static HtmlResourceSessionEntry CreateEntry(
        HtmlResourceKind kind,
        string source,
        string canonicalSource,
        HtmlResolvedResource resource) {
        byte[] digest;
        using (SHA256 sha = SHA256.Create()) digest = sha.ComputeHash(resource.EncodedBytes);
        string digestText = BitConverter.ToString(digest).Replace("-", string.Empty).ToLowerInvariant();
        return new HtmlResourceSessionEntry(kind, source, canonicalSource, resource.ContentType, resource.Length, digestText);
    }

    private void ValidateLimits() {
        if (ResourceTimeout <= TimeSpan.Zero || ResourceTimeout == Timeout.InfiniteTimeSpan) throw new ArgumentOutOfRangeException(nameof(ResourceTimeout));
        if (MaxConcurrentLoads <= 0) throw new ArgumentOutOfRangeException(nameof(MaxConcurrentLoads));
        if (MaxResourceBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxResourceBytes));
        if (MaxTotalResourceBytes < MaxResourceBytes) throw new ArgumentOutOfRangeException(nameof(MaxTotalResourceBytes));
        if (MaxResourceCount <= 0) throw new ArgumentOutOfRangeException(nameof(MaxResourceCount));
        if (MaxResourceRequests < MaxResourceCount) throw new ArgumentOutOfRangeException(nameof(MaxResourceRequests));
        if (MaxStylesheetImportDepth <= 0) throw new ArgumentOutOfRangeException(nameof(MaxStylesheetImportDepth));
    }

    private static bool IsAcceptedContentType(HtmlResourceKind kind, string contentType) {
        string normalized = contentType.Split(';')[0].Trim();
        if (kind == HtmlResourceKind.Image) return normalized.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
        if (kind == HtmlResourceKind.Font) {
            return normalized.StartsWith("font/", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("application/font-", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("application/x-font-", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "application/octet-stream", StringComparison.OrdinalIgnoreCase);
        }
        return kind == HtmlResourceKind.Stylesheet
            && (string.Equals(normalized, "text/css", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalized, "application/css", StringComparison.OrdinalIgnoreCase));
    }
}

/// <summary>Immutable accepted-resource identity and evidence.</summary>
public sealed class HtmlResourceSessionEntry {
    internal HtmlResourceSessionEntry(HtmlResourceKind kind, string source, string canonicalSource,
        string contentType, long length, string sha256) {
        Kind = kind;
        Source = source;
        CanonicalSource = canonicalSource;
        ContentType = contentType;
        Length = length;
        Sha256 = sha256;
    }

    /// <summary>Resource kind requested by the document.</summary>
    public HtmlResourceKind Kind { get; }
    /// <summary>Original policy-approved source.</summary>
    public string Source { get; }
    /// <summary>Canonical absolute URI or data source used as the cache identity.</summary>
    public string CanonicalSource { get; }
    /// <summary>Validated MIME type.</summary>
    public string ContentType { get; }
    /// <summary>Accepted encoded byte length.</summary>
    public long Length { get; }
    /// <summary>Lower-case SHA-256 content digest.</summary>
    public string Sha256 { get; }
}

internal static class HtmlRenderResourceLoader {
    private const string ComponentName = "OfficeIMO.Html.Renderer";

    private readonly struct ResourceResolution {
        internal ResourceResolution(bool isHandled, HtmlResolvedResource? resource) {
            IsHandled = isHandled;
            Resource = resource;
        }

        internal bool IsHandled { get; }
        internal HtmlResolvedResource? Resource { get; }
    }

    private delegate Task<ResourceResolution> ResourceResolver(
        HtmlRenderResourceRequest request,
        CancellationToken cancellationToken);

    private readonly struct PendingResource {
        internal PendingResource(HtmlResourceReference reference, int importDepth) {
            Reference = reference;
            ImportDepth = importDepth;
        }

        internal HtmlResourceReference Reference { get; }
        internal int ImportDepth { get; }
    }

    private readonly struct CompletedResolution {
        internal CompletedResolution(PendingResource pending, Uri uri, ResourceResolution resolution, Exception? exception) {
            Pending = pending;
            Uri = uri;
            Resolution = resolution;
            Exception = exception;
        }

        internal PendingResource Pending { get; }
        internal Uri Uri { get; }
        internal ResourceResolution Resolution { get; }
        internal Exception? Exception { get; }
    }

    internal static HtmlResourceSession Load(
        HtmlResourceManifest manifest,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics,
        CancellationToken cancellationToken,
        HtmlCssByteBudget? cssBudget = null) {
        var session = new HtmlResourceSession(options, diagnostics);
        HtmlRenderSynchronousResourceResolver? resolver = session.SynchronousResolver;
        if (resolver == null) return session;
        return LoadCoreAsync(
                manifest,
                options,
                session,
                cancellationToken,
                cssBudget,
                markAttemptedBeforeResolve: false,
                resolver: (request, token) => {
                    bool isHandled = resolver(
                        request,
                        token,
                        out HtmlResolvedResource? resource);
                    return Task.FromResult(
                        new ResourceResolution(isHandled, resource));
                })
            .GetAwaiter()
            .GetResult();
    }

    internal static Task<HtmlResourceSession> LoadAsync(
        HtmlResourceManifest manifest,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics,
        CancellationToken cancellationToken,
        HtmlCssByteBudget? cssBudget = null) {
        var session = new HtmlResourceSession(options, diagnostics);
        HtmlRenderResourceResolver? resolver = session.Resolver;
        if (resolver == null) {
            return Task.FromResult(session);
        }
        return LoadCoreAsync(
            manifest,
            options,
            session,
            cancellationToken,
            cssBudget,
            markAttemptedBeforeResolve: true,
            resolver: async (request, token) => new ResourceResolution(
                true,
                await resolver(request, token).ConfigureAwait(false)));
    }

    private static async Task<HtmlResourceSession> LoadCoreAsync(
        HtmlResourceManifest manifest,
        HtmlRenderOptions options,
        HtmlResourceSession result,
        CancellationToken cancellationToken,
        HtmlCssByteBudget? cssBudget,
        bool markAttemptedBeforeResolve,
        ResourceResolver resolver) {
        HtmlDiagnosticReport diagnostics = result.Diagnostics;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pending = new Queue<PendingResource>();
        foreach (HtmlResourceReference reference in manifest.Resources) {
            pending.Enqueue(new PendingResource(reference, 0));
        }

        var resourceOptions = new HtmlResourcePipelineOptions {
            ResourceUrlPolicy = result.ResourcePolicy.Clone(),
            MaxResponsiveImageCandidates = options.ResponsiveImageCandidateLimit,
            MediaContext = options.MediaContext,
            MediaWidth = options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth,
            MediaHeight = options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D
        };
        bool stop = false;
        int concurrency = markAttemptedBeforeResolve ? result.MaxConcurrentLoads : 1;
        while (pending.Count > 0 && !stop) {
            cancellationToken.ThrowIfCancellationRequested();
            if (result.AcceptedResourceCount >= result.MaxResourceCount) {
                diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded, "Resolved resources exceeded the configured operation-wide count limit.", HtmlDiagnosticSeverity.Error, detail: "limit=" + result.MaxResourceCount, lossKind: HtmlConversionLossKind.Omission);
                break;
            }

            int batchCapacity = Math.Min(concurrency, result.MaxResourceCount - result.AcceptedResourceCount);
            var tasks = new List<Task<CompletedResolution>>(batchCapacity);
            while (tasks.Count < batchCapacity && pending.Count > 0) {
                PendingResource pendingResource = pending.Dequeue();
                HtmlResourceReference reference = pendingResource.Reference;
                if (!reference.IsAllowed || !IsLoadableKind(reference.Kind) || reference.ResolvedSource.Length == 0) continue;
                if (reference.ResolvedSource.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) continue;
                if (!seen.Add(reference.ResolvedSource)) continue;
                if (!result.TryReserveRequest(reference)) {
                    stop = true;
                    break;
                }

                if (!Uri.TryCreate(reference.ResolvedSource, UriKind.Absolute, out Uri? uri)) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceUriInvalid, "A policy-approved resource could not be represented as an absolute URI.", HtmlDiagnosticSeverity.Warning, reference.Source, reference.ResolvedSource, HtmlConversionLossKind.Omission);
                    continue;
                }

                if (markAttemptedBeforeResolve) {
                    result.MarkAttempted(reference);
                }

                tasks.Add(ResolvePendingAsync(pendingResource, uri, result.ResourceTimeout, resolver, cancellationToken));
            }

            if (tasks.Count == 0) continue;
            CompletedResolution[] completed = await Task.WhenAll(tasks).ConfigureAwait(false);
            foreach (CompletedResolution item in completed) {
                cancellationToken.ThrowIfCancellationRequested();
                PendingResource pendingResource = item.Pending;
                HtmlResourceReference reference = pendingResource.Reference;
                if (item.Exception != null) {
                    if (item.Exception is HtmlRenderResourceByteLimitException byteLimit) {
                        result.MarkAttempted(reference);
                        diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded, "A resolved resource exceeded the configured per-resource byte limit.", HtmlDiagnosticSeverity.Warning, reference.Source, "bytes=" + byteLimit.ActualBytes, HtmlConversionLossKind.Omission);
                    } else if (item.Exception is OperationCanceledException) {
                        diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceTimeout, "Resource resolution exceeded the configured timeout.", HtmlDiagnosticSeverity.Warning, reference.Source, reference.ResolvedSource, HtmlConversionLossKind.Omission);
                    } else {
                        diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceLoadFailed, "The configured resource resolver failed to load a resource.", HtmlDiagnosticSeverity.Warning, reference.Source, item.Exception.GetType().Name, HtmlConversionLossKind.Omission);
                    }
                    continue;
                }

                ResourceResolution resolution = item.Resolution;
                if (!resolution.IsHandled) {
                    continue;
                }

                if (!markAttemptedBeforeResolve) {
                    result.MarkAttempted(reference);
                }
                HtmlResolvedResource? resource = resolution.Resource;
                if (resource == null) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceUnavailable, "The configured resource resolver did not return content.", HtmlDiagnosticSeverity.Warning, reference.Source, reference.ResolvedSource, HtmlConversionLossKind.Omission);
                    continue;
                }

                if (!result.TryAccept(reference, resource, out bool stopAfterResource)) {
                    if (stopAfterResource) stop = true;
                    continue;
                }
                if (reference.Kind == HtmlResourceKind.Stylesheet
                    && HtmlRenderStylesheetText.TryDecode(resource.EncodedBytes, out string css)) {
                    if (cssBudget != null
                        && !HtmlRenderStylesheetApplier.TryReserveCss(cssBudget, css, reference.Source, diagnostics)) {
                        result.MarkStylesheetRejected(reference);
                        continue;
                    }

                    if (cssBudget != null) result.MarkStylesheetBudgeted(reference);
                    EnqueueStylesheetResources(
                        pending,
                        css,
                        item.Uri,
                        pendingResource.ImportDepth,
                        resourceOptions,
                        result.MaxStylesheetImportDepth,
                        diagnostics);
                }
            }
        }

        return result;
    }

    private static async Task<CompletedResolution> ResolvePendingAsync(
        PendingResource pending,
        Uri uri,
        TimeSpan resourceTimeout,
        ResourceResolver resolver,
        CancellationToken cancellationToken) {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(resourceTimeout);
        try {
            HtmlResourceReference reference = pending.Reference;
            var request = new HtmlRenderResourceRequest(uri, reference.Source, reference.Kind);
            ResourceResolution resolution = await resolver(request, timeout.Token).ConfigureAwait(false);
            return new CompletedResolution(pending, uri, resolution, null);
        } catch (Exception exception) {
            return new CompletedResolution(pending, uri, default, exception);
        }
    }

    private static void EnqueueStylesheetResources(
        Queue<PendingResource> pending,
        string css,
        Uri stylesheetUri,
        int importDepth,
        HtmlResourcePipelineOptions resourceOptions,
        int maxStylesheetImportDepth,
        HtmlDiagnosticReport diagnostics) {
        HtmlExternalStylesheetAnalysis analysis = HtmlResourcePipeline.AnalyzeExternalStylesheet(css, stylesheetUri, resourceOptions);
        foreach (HtmlResourceReference imageResource in analysis.ImageResources) {
            if (imageResource.IsAllowed) {
                pending.Enqueue(new PendingResource(imageResource, importDepth));
            } else {
                diagnostics.Add(
                    ComponentName,
                    imageResource.DiagnosticCode.Length == 0 ? "ImageResourceRejectedByPolicy" : imageResource.DiagnosticCode,
                    "A stylesheet image source was rejected by the configured URL policy.",
                    HtmlDiagnosticSeverity.Warning,
                    imageResource.Source,
                    stylesheetUri.AbsoluteUri);
            }
        }

        foreach (HtmlResourceReference fontResource in analysis.FontResources) {
            if (fontResource.IsAllowed) {
                pending.Enqueue(new PendingResource(fontResource, importDepth));
            } else {
                diagnostics.Add(
                    ComponentName,
                    fontResource.DiagnosticCode.Length == 0 ? "FontResourceRejectedByPolicy" : fontResource.DiagnosticCode,
                    "A stylesheet font source was rejected by the configured URL policy.",
                    HtmlDiagnosticSeverity.Warning,
                    fontResource.Source,
                    stylesheetUri.AbsoluteUri);
            }
        }

        foreach (HtmlExternalStylesheetImport import in analysis.Imports) {
            if (!import.IsApplicable) {
                continue;
            }

            HtmlResourceReference reference = import.Reference;
            if (!reference.IsAllowed) {
                diagnostics.Add(
                    ComponentName,
                    reference.DiagnosticCode.Length == 0 ? "StylesheetResourceRejectedByPolicy" : reference.DiagnosticCode,
                    "A stylesheet import was rejected by the configured URL policy.",
                    HtmlDiagnosticSeverity.Warning,
                    reference.Source,
                    stylesheetUri.AbsoluteUri);
                continue;
            }

            if (importDepth >= maxStylesheetImportDepth) {
                diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.StylesheetImportDepthExceeded,
                    "Stylesheet imports exceeded the configured recursion depth.",
                    HtmlDiagnosticSeverity.Error,
                    reference.Source,
                    "limit=" + maxStylesheetImportDepth);
                continue;
            }

            pending.Enqueue(new PendingResource(reference, importDepth + 1));
        }
    }

    private static bool IsLoadableKind(HtmlResourceKind kind) =>
        kind == HtmlResourceKind.Image || kind == HtmlResourceKind.Stylesheet || kind == HtmlResourceKind.Font;

}
