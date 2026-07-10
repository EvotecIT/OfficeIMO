namespace OfficeIMO.Html;

internal sealed class HtmlRenderResourceSet {
    private readonly Dictionary<string, HtmlResolvedResource> _resources = new Dictionary<string, HtmlResolvedResource>(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _attempted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    internal void MarkAttempted(HtmlResourceReference reference) {
        if (reference.Source.Length > 0) _attempted.Add(reference.Source);
        if (reference.ResolvedSource.Length > 0) _attempted.Add(reference.ResolvedSource);
    }

    internal void Add(HtmlResourceReference reference, HtmlResolvedResource resource) {
        if (reference.Source.Length > 0) _resources[reference.Source] = resource;
        if (reference.ResolvedSource.Length > 0) _resources[reference.ResolvedSource] = resource;
    }

    internal bool TryGet(string? source, string? resolvedSource, out HtmlResolvedResource resource) {
        if (!string.IsNullOrWhiteSpace(source) && _resources.TryGetValue(source!, out resource!)) return true;
        if (!string.IsNullOrWhiteSpace(resolvedSource) && _resources.TryGetValue(resolvedSource!, out resource!)) return true;
        resource = null!;
        return false;
    }

    internal bool WasAttempted(string? source, string? resolvedSource) =>
        !string.IsNullOrWhiteSpace(source) && _attempted.Contains(source!)
        || !string.IsNullOrWhiteSpace(resolvedSource) && _attempted.Contains(resolvedSource!);
}

internal static class HtmlRenderResourceLoader {
    private const string ComponentName = "OfficeIMO.Html.Renderer";

    internal static async Task<HtmlRenderResourceSet> LoadAsync(HtmlResourceManifest manifest, HtmlRenderOptions options, HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken) {
        var result = new HtmlRenderResourceSet();
        if (options.ResourceResolver == null) return result;
        long totalBytes = 0L;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (HtmlResourceReference reference in manifest.Resources) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!reference.IsAllowed || reference.Kind != HtmlResourceKind.Image || reference.ResolvedSource.Length == 0) continue;
            if (reference.ResolvedSource.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) continue;
            if (!seen.Add(reference.ResolvedSource)) continue;
            result.MarkAttempted(reference);

            if (!Uri.TryCreate(reference.ResolvedSource, UriKind.Absolute, out Uri? uri)) {
                diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceUriInvalid, "A policy-approved resource could not be represented as an absolute URI.", HtmlDiagnosticSeverity.Warning, reference.Source, reference.ResolvedSource);
                continue;
            }

            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(options.ResourceTimeout);
            try {
                var request = new HtmlRenderResourceRequest(uri, reference.Source, reference.Kind);
                HtmlResolvedResource? resource = await options.ResourceResolver(request, timeout.Token).ConfigureAwait(false);
                if (resource == null) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceUnavailable, "The configured resource resolver did not return content.", HtmlDiagnosticSeverity.Warning, reference.Source, reference.ResolvedSource);
                    continue;
                }

                long length = resource.Bytes.LongLength;
                if (length > options.MaxResourceBytes) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded, "A resolved resource exceeded the configured per-resource byte limit.", HtmlDiagnosticSeverity.Warning, reference.Source, "bytes=" + length);
                    continue;
                }

                if (totalBytes + length > options.MaxTotalResourceBytes) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded, "Resolved resources exceeded the configured total byte limit.", HtmlDiagnosticSeverity.Error, reference.Source, "bytes=" + (totalBytes + length));
                    break;
                }

                if (!resource.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) {
                    diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceContentTypeRejected, "An image resolver returned a non-image media type.", HtmlDiagnosticSeverity.Warning, reference.Source, resource.ContentType);
                    continue;
                }

                totalBytes += length;
                result.Add(reference, resource);
            } catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested) {
                diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceTimeout, "Resource resolution exceeded the configured timeout.", HtmlDiagnosticSeverity.Warning, reference.Source, reference.ResolvedSource);
            } catch (Exception exception) {
                diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ResourceLoadFailed, "The configured resource resolver failed to load a resource.", HtmlDiagnosticSeverity.Warning, reference.Source, exception.GetType().Name);
            }
        }

        return result;
    }
}
