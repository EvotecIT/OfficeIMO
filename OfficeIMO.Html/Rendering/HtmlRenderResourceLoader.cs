namespace OfficeIMO.Html;

internal sealed class HtmlRenderResourceSet {
    private readonly Dictionary<string, HtmlResolvedResource> _resources = new Dictionary<string, HtmlResolvedResource>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _resolvedSources = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _attempted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    internal long AcceptedResourceBytes { get; private set; }
    internal int AcceptedResourceCount { get; private set; }

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
    }

    internal bool CanAcceptInlineResource(
        long estimatedBytes,
        HtmlRenderOptions options,
        out string diagnosticCode,
        out string diagnosticDetail) {
        if (estimatedBytes > options.MaxResourceBytes) {
            diagnosticCode = HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded;
            diagnosticDetail = "bytes=" + estimatedBytes;
            return false;
        }

        if (AcceptedResourceCount >= options.MaxResourceCount) {
            diagnosticCode = HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded;
            diagnosticDetail = "limit=" + options.MaxResourceCount;
            return false;
        }

        if (estimatedBytes > options.MaxTotalResourceBytes - AcceptedResourceBytes) {
            diagnosticCode = HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded;
            diagnosticDetail = "bytes=" + (AcceptedResourceBytes + estimatedBytes);
            return false;
        }

        diagnosticCode = string.Empty;
        diagnosticDetail = string.Empty;
        return true;
    }

    internal void AddInline(string resolvedSource, HtmlResolvedResource resource) {
        if (string.IsNullOrWhiteSpace(resolvedSource) || _resources.ContainsKey(resolvedSource)) return;
        AcceptedResourceBytes += resource.Length;
        AcceptedResourceCount++;
        _resources[resolvedSource] = resource;
        _resolvedSources[resolvedSource] = resolvedSource;
    }

    internal bool TryGet(string? source, string? resolvedSource, out HtmlResolvedResource resource) {
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

    internal static HtmlRenderResourceSet Load(
        HtmlResourceManifest manifest,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics,
        CancellationToken cancellationToken) {
        HtmlRenderSynchronousResourceResolver? resolver =
            options.SynchronousResourceResolver;
        if (resolver == null) return new HtmlRenderResourceSet();
        return LoadCoreAsync(
                manifest,
                options,
                diagnostics,
                cancellationToken,
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

    internal static Task<HtmlRenderResourceSet> LoadAsync(
        HtmlResourceManifest manifest,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics,
        CancellationToken cancellationToken) {
        HtmlRenderResourceResolver? resolver = options.ResourceResolver;
        if (resolver == null) {
            return Task.FromResult(new HtmlRenderResourceSet());
        }
        return LoadCoreAsync(
            manifest,
            options,
            diagnostics,
            cancellationToken,
            markAttemptedBeforeResolve: true,
            resolver: async (request, token) => new ResourceResolution(
                true,
                await resolver(request, token).ConfigureAwait(false)));
    }

    private static async Task<HtmlRenderResourceSet> LoadCoreAsync(
        HtmlResourceManifest manifest,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics,
        CancellationToken cancellationToken,
        bool markAttemptedBeforeResolve,
        ResourceResolver resolver) {
        var result = new HtmlRenderResourceSet();
        long totalBytes = 0L;
        int resourceCount = 0;
        int requestCount = 0;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pending = new Queue<PendingResource>();
        foreach (HtmlResourceReference reference in manifest.Resources) {
            pending.Enqueue(new PendingResource(reference, 0));
        }

        var resourceOptions = new HtmlResourcePipelineOptions {
            ResourceUrlPolicy = options.GetResourceUrlPolicy().Clone(),
            MaxResponsiveImageCandidates = options.ResponsiveImageCandidateLimit,
            MediaContext = options.MediaContext,
            MediaWidth = options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth,
            MediaHeight = options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D
        };
        bool stop = false;
        int concurrency = markAttemptedBeforeResolve ? options.MaxConcurrentResourceLoads : 1;
        while (pending.Count > 0 && !stop) {
            cancellationToken.ThrowIfCancellationRequested();
            if (resourceCount >= options.MaxResourceCount) {
                diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded, "Resolved resources exceeded the configured operation-wide count limit.", HtmlDiagnosticSeverity.Error, detail: "limit=" + options.MaxResourceCount, lossKind: HtmlConversionLossKind.Omission);
                break;
            }

            int batchCapacity = Math.Min(concurrency, options.MaxResourceCount - resourceCount);
            var tasks = new List<Task<CompletedResolution>>(batchCapacity);
            while (tasks.Count < batchCapacity && pending.Count > 0) {
                PendingResource pendingResource = pending.Dequeue();
                HtmlResourceReference reference = pendingResource.Reference;
                if (!reference.IsAllowed || !IsLoadableKind(reference.Kind) || reference.ResolvedSource.Length == 0) continue;
                if (reference.ResolvedSource.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) continue;
                if (!seen.Add(reference.ResolvedSource)) continue;
                if (requestCount >= options.MaxResourceRequests) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceRequestLimitExceeded, "Resource resolver invocations exceeded the configured operation-wide request limit.", HtmlDiagnosticSeverity.Error, reference.Source, "limit=" + options.MaxResourceRequests, HtmlConversionLossKind.Omission);
                    stop = true;
                    break;
                }

                if (!Uri.TryCreate(reference.ResolvedSource, UriKind.Absolute, out Uri? uri)) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceUriInvalid, "A policy-approved resource could not be represented as an absolute URI.", HtmlDiagnosticSeverity.Warning, reference.Source, reference.ResolvedSource, HtmlConversionLossKind.Omission);
                    continue;
                }

                requestCount++;
                if (markAttemptedBeforeResolve) {
                    result.MarkAttempted(reference);
                }

                tasks.Add(ResolvePendingAsync(pendingResource, uri, options.ResourceTimeout, resolver, cancellationToken));
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

                long length = resource.Length;
                if (length > options.MaxResourceBytes) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded, "A resolved resource exceeded the configured per-resource byte limit.", HtmlDiagnosticSeverity.Warning, reference.Source, "bytes=" + length, HtmlConversionLossKind.Omission);
                    continue;
                }

                if (totalBytes + length > options.MaxTotalResourceBytes) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded, "Resolved resources exceeded the configured total byte limit.", HtmlDiagnosticSeverity.Error, reference.Source, "bytes=" + (totalBytes + length), HtmlConversionLossKind.Omission);
                    stop = true;
                    break;
                }

                if (!IsAcceptedContentType(reference.Kind, resource.ContentType)) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceContentTypeRejected, "A resolver returned an incompatible media type for the requested resource kind.", HtmlDiagnosticSeverity.Warning, reference.Source, reference.Kind + ":" + resource.ContentType, HtmlConversionLossKind.Omission);
                    continue;
                }

                totalBytes += length;
                resourceCount++;
                result.Add(reference, resource);
                if (reference.Kind == HtmlResourceKind.Stylesheet
                    && HtmlRenderStylesheetText.TryDecode(resource.EncodedBytes, out string css)) {
                    EnqueueStylesheetResources(
                        pending,
                        css,
                        item.Uri,
                        pendingResource.ImportDepth,
                        resourceOptions,
                        options,
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
        HtmlRenderOptions options,
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

            if (importDepth >= options.MaxStylesheetImportDepth) {
                diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.StylesheetImportDepthExceeded,
                    "Stylesheet imports exceeded the configured recursion depth.",
                    HtmlDiagnosticSeverity.Error,
                    reference.Source,
                    "limit=" + options.MaxStylesheetImportDepth);
                continue;
            }

            pending.Enqueue(new PendingResource(reference, importDepth + 1));
        }
    }

    private static bool IsLoadableKind(HtmlResourceKind kind) =>
        kind == HtmlResourceKind.Image || kind == HtmlResourceKind.Stylesheet || kind == HtmlResourceKind.Font;

    private static bool IsAcceptedContentType(HtmlResourceKind kind, string contentType) {
        string normalized = contentType.Split(';')[0].Trim();
        if (kind == HtmlResourceKind.Image) {
            return normalized.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
        }

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
