using OfficeIMO.Html;

namespace OfficeIMO.Mhtml;

public sealed partial class MhtmlDocument {
    /// <summary>
    /// Creates an independent HTML conversion document whose direct image sources use the
    /// matching MIME parts from this archive. It never fetches resources outside the archive.
    /// The original <see cref="HtmlDocument"/> and its render resolver remain unchanged.
    /// </summary>
    /// <param name="resourceLimits">Optional shared resource byte, count, and request limits. Other render settings are ignored.</param>
    /// <param name="cancellationToken">Stops preparation before the next image is processed.</param>
    /// <returns>The editable document and diagnostics for image sources not embedded.</returns>
    public MhtmlImageEmbeddingResult CreateEmbeddedImageDocumentResult(
        HtmlRenderOptions? resourceLimits = null,
        CancellationToken cancellationToken = default) {
        HtmlRenderOptions limits = resourceLimits?.Clone() ?? new HtmlRenderOptions();
        if (limits.MaxResourceBytes <= 0) throw new ArgumentOutOfRangeException(nameof(resourceLimits), "The per-resource byte limit must be positive.");
        if (limits.MaxTotalResourceBytes <= 0) throw new ArgumentOutOfRangeException(nameof(resourceLimits), "The total resource byte limit must be positive.");
        if (limits.MaxResourceCount <= 0) throw new ArgumentOutOfRangeException(nameof(resourceLimits), "The resource count limit must be positive.");
        if (limits.MaxResourceRequests <= 0) throw new ArgumentOutOfRangeException(nameof(resourceLimits), "The resource request limit must be positive.");

        var diagnostics = new List<HtmlDiagnostic>();
        var embedded = new Dictionary<string, string?>(HtmlResourceIdentityComparer.Instance);
        var embeddedLengths = new Dictionary<string, long>(HtmlResourceIdentityComparer.Instance);
        var counted = new HashSet<string>(HtmlResourceIdentityComparer.Instance);
        int requests = 0;
        int resourceCount = 0;
        long resourceBytes = 0;
        HtmlConversionDocument editable = HtmlDocument.Edit(document => {
            // The editable HTML repeats a data URI for every img element, even when all of
            // them share one MIME part. Bound that expansion against the source contract.
            long projectedChars = document.OuterHtml.Length;
            foreach (var image in document.QuerySelectorAll("img[src]")) {
                cancellationToken.ThrowIfCancellationRequested();
                string source = image.GetAttribute("src")?.Trim() ?? string.Empty;
                if (source.Length == 0 || source.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) continue;

                // Resolve the authored src itself: a responsive <picture> may select a
                // <source> and omit its fallback img[src] from the resource manifest.
                string resolvedSource = HtmlUrlPolicyEvaluator.ResolveUrl(
                    source, HtmlDocument.BaseUri ?? BaseUri, _imageResourceUrlPolicy);
                if (!Uri.TryCreate(resolvedSource, UriKind.Absolute, out Uri? uri)) {
                    AddImageDiagnostic(diagnostics, "ImageResourceRejectedByPolicy",
                        "The image source was rejected by the HTML resource URL policy.", source);
                    continue;
                }
                if (!embedded.TryGetValue(resolvedSource, out string? dataUri)) {
                    if (++requests > limits.MaxResourceRequests) {
                        AddImageDiagnostic(diagnostics, HtmlRenderDiagnosticCodes.ResourceRequestLimitExceeded,
                            "The image source exceeded the resource request limit.", source);
                        embedded.Add(resolvedSource, null);
                        continue;
                    }
                    MhtmlResource? resource = FindResource(source, uri);
                    if (resource == null || resource.Length == 0) {
                        AddImageDiagnostic(diagnostics, HtmlRenderDiagnosticCodes.ResourceUnavailable,
                            "The image source has no matching nonempty MIME part in the archive.", source);
                        embedded.Add(resolvedSource, null);
                        continue;
                    }
                    string mediaType = resource.ContentType.Split(';')[0].Trim().ToLowerInvariant();
                    if (!mediaType.StartsWith("image/", StringComparison.Ordinal)) {
                        AddImageDiagnostic(diagnostics, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
                            "The archived image source has a non-image media type.", source, mediaType);
                        embedded.Add(resolvedSource, null);
                        continue;
                    }
                    if (resource.Length > limits.MaxResourceBytes) {
                        AddImageDiagnostic(diagnostics, HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded,
                            "The archived image exceeded the per-resource byte limit.", source,
                            $"bytes={resource.Length}; limit={limits.MaxResourceBytes}");
                        embedded.Add(resolvedSource, null);
                        continue;
                    }
                    dataUri = "data:" + mediaType + ";base64," + Convert.ToBase64String(resource.EncodedContent);
                    embedded.Add(resolvedSource, dataUri);
                    embeddedLengths.Add(resolvedSource, resource.Length);
                }
                if (dataUri == null) continue;
                long growth = Math.Max(0L, (long)dataUri.Length - source.Length);
                if (_editableSourceCharacterLimit is int maximum
                    && growth > (long)maximum - projectedChars - 512) {
                    AddImageDiagnostic(diagnostics, HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded,
                        "Embedding the image would exceed the editable HTML source limit.", source,
                        $"projected={projectedChars + growth}; limit={maximum}");
                    continue;
                }
                if (!counted.Contains(resolvedSource)) {
                    long length = embeddedLengths[resolvedSource];
                    if (resourceCount >= limits.MaxResourceCount) {
                        AddImageDiagnostic(diagnostics, HtmlRenderDiagnosticCodes.ResourceCountLimitExceeded,
                            "The archived image exceeded the resource count limit.", source);
                        embedded[resolvedSource] = null;
                        continue;
                    }
                    if (length > limits.MaxTotalResourceBytes - resourceBytes) {
                        AddImageDiagnostic(diagnostics, HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded,
                            "The archived image exceeded the total resource byte limit.", source,
                            $"bytes={length}; remaining={limits.MaxTotalResourceBytes - resourceBytes}");
                        embedded[resolvedSource] = null;
                        continue;
                    }
                    counted.Add(resolvedSource);
                    resourceCount++;
                    resourceBytes += length;
                }
                image.SetAttribute("src", dataUri);
                projectedChars += growth;
            }
        });
        return new MhtmlImageEmbeddingResult(editable, diagnostics, resourceCount, resourceBytes);
    }

    private static void AddImageDiagnostic(List<HtmlDiagnostic> diagnostics, string code,
        string message, string source, string? detail = null) {
        diagnostics.Add(new HtmlDiagnostic("OfficeIMO.Mhtml", code, message,
            HtmlDiagnosticSeverity.Warning, source, detail, OfficeConversionLossKind.Omission));
    }
}
