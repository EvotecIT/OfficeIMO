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
        AcceptedResourceBytes += resource.Bytes.LongLength;
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
        AcceptedResourceBytes += resource.Bytes.LongLength;
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
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pending = new Queue<PendingResource>();
        foreach (HtmlResourceReference reference in manifest.Resources) {
            pending.Enqueue(new PendingResource(reference, 0));
        }

        var resourceOptions = new HtmlResourcePipelineOptions {
            UrlPolicy = options.GetResourceUrlPolicy().Clone(),
            MediaContext = options.MediaContext,
            MediaWidth = options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth,
            MediaHeight = options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D
        };
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            PendingResource pendingResource = pending.Dequeue();
            HtmlResourceReference reference = pendingResource.Reference;
            if (!reference.IsAllowed || !IsLoadableKind(reference.Kind) || reference.ResolvedSource.Length == 0) continue;
            if (reference.ResolvedSource.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) continue;
            if (!seen.Add(reference.ResolvedSource)) continue;
            if (resourceCount >= options.MaxResourceCount) {
                diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded, "Resolved resources exceeded the configured operation-wide count limit.", HtmlDiagnosticSeverity.Error, reference.Source, "limit=" + options.MaxResourceCount, HtmlConversionLossKind.Omission);
                break;
            }

            if (!Uri.TryCreate(reference.ResolvedSource, UriKind.Absolute, out Uri? uri)) {
                diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceUriInvalid, "A policy-approved resource could not be represented as an absolute URI.", HtmlDiagnosticSeverity.Warning, reference.Source, reference.ResolvedSource, HtmlConversionLossKind.Omission);
                continue;
            }

            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(options.ResourceTimeout);
            try {
                var request = new HtmlRenderResourceRequest(uri, reference.Source, reference.Kind);
                if (markAttemptedBeforeResolve) {
                    result.MarkAttempted(reference);
                }
                ResourceResolution resolution = await resolver(
                    request,
                    timeout.Token).ConfigureAwait(false);
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

                long length = resource.Bytes.LongLength;
                if (length > options.MaxResourceBytes) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded, "A resolved resource exceeded the configured per-resource byte limit.", HtmlDiagnosticSeverity.Warning, reference.Source, "bytes=" + length, HtmlConversionLossKind.Omission);
                    continue;
                }

                if (totalBytes + length > options.MaxTotalResourceBytes) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded, "Resolved resources exceeded the configured total byte limit.", HtmlDiagnosticSeverity.Error, reference.Source, "bytes=" + (totalBytes + length), HtmlConversionLossKind.Omission);
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
                    && HtmlRenderStylesheetText.TryDecode(resource.Bytes, out string css)) {
                    EnqueueStylesheetResources(
                        pending,
                        css,
                        uri,
                        pendingResource.ImportDepth,
                        resourceOptions,
                        options,
                        diagnostics);
                }
            } catch (HtmlRenderResourceByteLimitException exception) {
                result.MarkAttempted(reference);
                diagnostics.Add(
                    ComponentName,
                    HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded,
                    "A resolved resource exceeded the configured per-resource byte limit.",
                    HtmlDiagnosticSeverity.Warning,
                    reference.Source,
                    "bytes=" + exception.ActualBytes,
                    HtmlConversionLossKind.Omission);
            } catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested) {
                diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceTimeout, "Resource resolution exceeded the configured timeout.", HtmlDiagnosticSeverity.Warning, reference.Source, reference.ResolvedSource, HtmlConversionLossKind.Omission);
            } catch (Exception exception) {
                diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceLoadFailed, "The configured resource resolver failed to load a resource.", HtmlDiagnosticSeverity.Warning, reference.Source, exception.GetType().Name, HtmlConversionLossKind.Omission);
            }
        }

        return result;
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
