using OfficeIMO.Html;

namespace OfficeIMO.Epub.Image;

// The virtual publication origin and resource snapshots are shared across all chapters in one render.
internal sealed class EpubChapterResources {
    private readonly Dictionary<string, Lazy<HtmlResolvedResource?>> _resources = new(StringComparer.Ordinal);

    internal EpubChapterResources(EpubDocument source, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        foreach (EpubResource resource in source.Resources) {
            cancellationToken.ThrowIfCancellationRequested();
            if (resource.IsRemote || _resources.ContainsKey(resource.Path)) continue;
            _resources.Add(resource.Path, new Lazy<HtmlResolvedResource?>(() =>
                resource.Data is { Length: > 0 } bytes
                    ? HtmlResolvedResource.FromOwnedBytes(bytes, resource.MediaType ?? "application/octet-stream")
                    : null));
        }
    }

    internal HtmlConversionDocument PrepareDocument(EpubChapter chapter, HtmlRenderOptions options,
        bool includeChapterTitle, CancellationToken cancellationToken, bool allowEmbeddedResources = true, int? maximumHtmlNodes = null) {
        Uri baseUri = CreateChapterUri(chapter);
        options.BaseUri = baseUri;
        if (allowEmbeddedResources) Configure(options);
        var documentOptions = new HtmlConversionDocumentOptions {
            BaseUri = baseUri,
            UrlPolicy = options.UrlPolicy.Clone(),
            ResourceUrlPolicy = options.GetResourceUrlPolicy().Clone(),
            UseBodyContentsOnly = false
        };
        if (maximumHtmlNodes.HasValue) documentOptions.Limits.MaxHtmlNodes = maximumHtmlNodes;
        if (string.IsNullOrWhiteSpace(chapter.Html))
            return HtmlConversionDocument.Parse(CreatePlainTextChapter(chapter, includeChapterTitle), documentOptions, cancellationToken);
        return string.Equals(chapter.MediaType, "application/xhtml+xml", StringComparison.OrdinalIgnoreCase)
            ? HtmlConversionDocument.ParseXhtml(chapter.Html!, documentOptions, cancellationToken)
            : HtmlConversionDocument.Parse(chapter.Html!, documentOptions, cancellationToken);
    }

    private void Configure(HtmlRenderOptions options) {
        HtmlUrlPolicy original = options.GetResourceUrlPolicy().Clone();
        HtmlUrlPolicy policy = original.Clone();
        policy.RestrictUrlSchemes = true;
        policy.AllowedUrlSchemes.Add("epub");
        options.ResourceUrlPolicy = policy;

        HtmlRenderSynchronousResourceResolver? syncFallback = options.SynchronousResourceResolver;
        options.SynchronousResourceResolver = (
            HtmlRenderResourceRequest request, CancellationToken token, out HtmlResolvedResource? resource) => {
            token.ThrowIfCancellationRequested();
            if (TryResolve(request.Uri, options.MaxResourceBytes, out resource)) return true;
            return syncFallback != null && HtmlUrlPolicyEvaluator.IsAllowed(request.Uri.AbsoluteUri, original)
                && syncFallback(request, token, out resource);
        };
        HtmlRenderResourceResolver? fallback = options.ResourceResolver;
        options.ResourceResolver = (request, token) => {
            token.ThrowIfCancellationRequested();
            if (TryResolve(request.Uri, options.MaxResourceBytes, out HtmlResolvedResource? resource))
                return Task.FromResult(resource);
            return fallback != null && HtmlUrlPolicyEvaluator.IsAllowed(request.Uri.AbsoluteUri, original)
                ? fallback(request, token) : Task.FromResult<HtmlResolvedResource?>(null);
        };
    }

    private bool TryResolve(Uri uri, long maxResourceBytes, out HtmlResolvedResource? resource) {
        resource = null;
        if (!uri.Scheme.Equals("epub", StringComparison.OrdinalIgnoreCase)) return false;

        // A chapter may change its base URL, but it cannot change the publication's origin.
        if (!uri.Host.Equals("document", StringComparison.OrdinalIgnoreCase) || uri.Port != -1 || uri.UserInfo.Length != 0) return true;
        EpubReference reference = EpubReference.Resolve("package.opf", uri.AbsolutePath);
        if (reference.Kind == EpubReferenceKind.Container && reference.ContainerPath != null &&
            _resources.TryGetValue(reference.ContainerPath, out Lazy<HtmlResolvedResource?>? snapshot)) {
            resource = snapshot.Value;
            if (resource != null && resource.Length > maxResourceBytes)
                throw new HtmlRenderResourceByteLimitException(resource.Length);
        }
        return true;
    }

    private static string CreatePlainTextChapter(
        EpubChapter chapter,
        bool includeTitle) {
        var builder = new StringBuilder(
            "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>");
        if (includeTitle && !string.IsNullOrWhiteSpace(chapter.Title)) {
            builder.Append("<h1>")
                .Append(WebUtility.HtmlEncode(chapter.Title))
                .Append("</h1>");
        }
        builder.Append("<div style=\"white-space:pre-wrap\">")
            .Append(WebUtility.HtmlEncode(chapter.Text))
            .Append("</div></body></html>");
        return builder.ToString();
    }

    private static Uri CreateChapterUri(EpubChapter chapter) {
        string path = NormalizePath(chapter.Path);
        // The HTML parser applies the retained markup's base element once.
        return new Uri("epub://document/" + EscapePath(path));
    }

    private static string EscapePath(string path) =>
        string.Join(
            "/",
            path.Split('/')
                .Where(segment => segment.Length > 0)
                .Select(Uri.EscapeDataString));

    private static string NormalizePath(string path) {
        var segments = new List<string>();
        foreach (string segment in path
                     .Replace('\\', '/')
                     .Split('/')) {
            if (segment.Length == 0 || segment == ".") continue;
            if (segment == "..") {
                if (segments.Count > 0) segments.RemoveAt(
                    segments.Count - 1);
                continue;
            }
            segments.Add(segment);
        }
        return string.Join("/", segments);
    }

}
