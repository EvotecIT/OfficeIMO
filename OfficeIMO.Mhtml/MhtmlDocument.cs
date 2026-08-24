using OfficeIMO.Email;
using OfficeIMO.Html;

namespace OfficeIMO.Mhtml;

/// <summary>
/// Represents an MHTML web archive as an HTML document plus its decoded MIME related resources.
/// </summary>
public sealed class MhtmlDocument {
    private static readonly Uri FallbackBaseUri = new Uri("mhtml://archive/");
    private readonly EmailDocument _mimeDocument;
    private readonly IReadOnlyList<MhtmlResource> _resources;
    private readonly IReadOnlyList<EmailDiagnostic> _mimeDiagnostics;

    /// <summary>Creates an MHTML document from HTML and optional related resources.</summary>
    public MhtmlDocument(string html, IEnumerable<MhtmlResource>? resources = null,
        string? contentLocation = null, string? rootContentId = null, string? subject = null,
        HtmlConversionDocumentOptions? htmlOptions = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        _resources = (resources ?? Enumerable.Empty<MhtmlResource>()).ToArray();
        ContentLocation = NormalizeOptional(contentLocation);
        RootContentId = NormalizeContentId(rootContentId);
        Subject = NormalizeOptional(subject);
        BaseUri = ResolveBaseUri(ContentLocation, null);
        _mimeDiagnostics = BuildResourceDiagnostics(_resources, BaseUri, RootContentId, ContentLocation);
        HtmlDocument = HtmlConversionDocument.Parse(html, PrepareHtmlOptions(htmlOptions, BaseUri, _resources));
        _mimeDocument = CreateMimeDocument(html, _resources, ContentLocation, RootContentId, Subject);
    }

    private MhtmlDocument(EmailReadResult readResult, Uri? sourceBaseUri,
        HtmlConversionDocumentOptions? htmlOptions) {
        if (readResult == null) throw new ArgumentNullException(nameof(readResult));
        if (readResult.HasErrors) throw CreateReadException(readResult.Diagnostics);
        _mimeDocument = readResult.Document;
        string? html = _mimeDocument.Body.Html;
        if (html == null) throw new InvalidDataException("The MHTML archive does not contain an HTML root part.");
        if (!IsMultipartRelated(_mimeDocument.Headers)) {
            throw new InvalidDataException("The artifact is an RFC message but its root is not multipart/related MHTML.");
        }

        ContentLocation = NormalizeOptional(_mimeDocument.Body.HtmlContentLocation)
            ?? GetHeaderValue(_mimeDocument.Headers, "Snapshot-Content-Location")
            ?? GetHeaderValue(_mimeDocument.Headers, "Content-Location");
        RootContentId = NormalizeContentId(_mimeDocument.Body.HtmlContentId);
        Subject = NormalizeOptional(_mimeDocument.Subject);
        BaseUri = ResolveBaseUri(ContentLocation, sourceBaseUri);
        _resources = _mimeDocument.Attachments
            .Where(static attachment => attachment.IsMimeRelated)
            .Select(MhtmlResource.FromEmailAttachment)
            .ToArray();
        _mimeDiagnostics = readResult.Diagnostics
            .Concat(BuildResourceDiagnostics(_resources, BaseUri, RootContentId, ContentLocation))
            .ToArray();
        HtmlDocument = HtmlConversionDocument.Parse(html, PrepareHtmlOptions(htmlOptions, BaseUri, _resources));
    }

    /// <summary>Parsed HTML root document.</summary>
    public HtmlConversionDocument HtmlDocument { get; }

    /// <summary>Original HTML root source.</summary>
    public string Html => HtmlDocument.SourceHtml;

    /// <summary>Decoded related resources in archive order.</summary>
    public IReadOnlyList<MhtmlResource> Resources => _resources;

    /// <summary>Root content location, when declared.</summary>
    public string? ContentLocation { get; }

    /// <summary>Root Content-ID without angle brackets, when declared.</summary>
    public string? RootContentId { get; }

    /// <summary>Optional archive subject.</summary>
    public string? Subject { get; }

    /// <summary>Base URI used for HTML and related-resource resolution.</summary>
    public Uri BaseUri { get; }

    /// <summary>Diagnostics produced by the shared bounded MIME reader.</summary>
    public IReadOnlyList<EmailDiagnostic> MimeDiagnostics => _mimeDiagnostics;

    /// <summary>Loads an MHTML archive from a file.</summary>
    public static MhtmlDocument Load(string path, EmailReaderOptions? mimeOptions = null,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        EmailReadResult result = new EmailDocumentReader(mimeOptions ?? EmailReaderOptions.Default)
            .Read(path, cancellationToken);
        return new MhtmlDocument(result, CreateFileBaseUri(path), htmlOptions);
    }

    /// <summary>Loads an MHTML archive from a caller-owned stream.</summary>
    public static MhtmlDocument Load(Stream stream, EmailReaderOptions? mimeOptions = null,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        EmailReadResult result = new EmailDocumentReader(mimeOptions ?? EmailReaderOptions.Default)
            .Read(stream, cancellationToken);
        return new MhtmlDocument(result, null, htmlOptions);
    }

    /// <summary>Asynchronously loads an MHTML archive from a file.</summary>
    public static async Task<MhtmlDocument> LoadAsync(string path, EmailReaderOptions? mimeOptions = null,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        EmailReadResult result = await new EmailDocumentReader(mimeOptions ?? EmailReaderOptions.Default)
            .ReadAsync(path, cancellationToken).ConfigureAwait(false);
        return new MhtmlDocument(result, CreateFileBaseUri(path), htmlOptions);
    }

    /// <summary>Asynchronously loads an MHTML archive from a caller-owned stream.</summary>
    public static async Task<MhtmlDocument> LoadAsync(Stream stream, EmailReaderOptions? mimeOptions = null,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        EmailReadResult result = await new EmailDocumentReader(mimeOptions ?? EmailReaderOptions.Default)
            .ReadAsync(stream, cancellationToken).ConfigureAwait(false);
        return new MhtmlDocument(result, null, htmlOptions);
    }

    /// <summary>Creates a resolver that serves only resources embedded in this archive.</summary>
    public HtmlRenderResourceResolver CreateResourceResolver() => ResolveResourceAsync;

    /// <summary>
    /// Applies the archive base URI, resource-only URL policy, and embedded-resource resolver to render options.
    /// The hyperlink policy is left unchanged. Remote resources absent from the archive are fetched only
    /// through the policy's one-hop resolver so every redirect destination is approved before it is requested.
    /// </summary>
    public void ConfigureRenderOptions(HtmlRenderOptions options) => ConfigureRenderOptions(options, null);

    /// <summary>
    /// Applies the archive base URI, resource-only URL policy, embedded-resource resolver, and explicit
    /// bounded remote-resource policy to render options.
    /// </summary>
    public void ConfigureRenderOptions(HtmlRenderOptions options,
        MhtmlRemoteResourcePolicy? remoteResourcePolicy) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        remoteResourcePolicy ??= MhtmlRemoteResourcePolicy.CreateEmbeddedOnlyProfile();
        remoteResourcePolicy.ApplyLimits(options);
        options.BaseUri ??= BaseUri;
        options.UrlPolicy ??= HtmlUrlPolicy.CreateOfficeIMOProfile();
        HtmlUrlPolicy fallbackResourceUrlPolicy = (options.ResourceUrlPolicy ?? options.UrlPolicy).Clone();
        HtmlUrlPolicy resourceUrlPolicy = fallbackResourceUrlPolicy.Clone();
        if (resourceUrlPolicy.RestrictUrlSchemes) {
            resourceUrlPolicy.AllowedUrlSchemes.Add("cid");
            resourceUrlPolicy.AllowedUrlSchemes.Add(BaseUri.Scheme);
        }
        resourceUrlPolicy.DisallowFileUrls = false;
        options.ResourceUrlPolicy = resourceUrlPolicy;
        options.ResourceResolver = new ManagedResourceResolver(
            this,
            fallbackResourceUrlPolicy,
            remoteResourcePolicy,
            includeEmbeddedResources: true).ResolveAsync;
    }

    /// <summary>
    /// Reuses only a resolver created by this archive while selecting whether embedded MIME resources remain visible.
    /// This lets optional bridge packages retain the policy-owned one-hop remote resolver without trusting arbitrary delegates.
    /// </summary>
    internal bool TryReconfigureOwnedResourceResolver(
        HtmlRenderResourceResolver? resolver,
        bool includeEmbeddedResources,
        out HtmlRenderResourceResolver? configuredResolver) {
        if (resolver?.Target is ManagedResourceResolver managed
            && ReferenceEquals(managed.Document, this)) {
            configuredResolver = managed.WithEmbeddedResolution(includeEmbeddedResources).ResolveAsync;
            return true;
        }

        configuredResolver = null;
        return false;
    }

    private async Task<HtmlResolvedResource?> ResolveRemoteResourceAsync(
        HtmlRenderResourceRequest request,
        HtmlUrlPolicy resourceUrlPolicy,
        MhtmlRemoteResourcePolicy remoteResourcePolicy,
        CancellationToken cancellationToken) {
        Uri current = request.Uri;
        for (int redirectNumber = 0; ; redirectNumber++) {
            string approvedSource = HtmlUrlPolicyEvaluator.ResolveUrl(
                current.AbsoluteUri,
                baseUri: null,
                resourceUrlPolicy);
            if (!Uri.TryCreate(approvedSource, UriKind.Absolute, out Uri? approvedUri)
                || !approvedUri.Equals(current)
                || !remoteResourcePolicy.AllowsRequest(current, BaseUri)) return null;
            if (redirectNumber > 0 && !request.TryReserveAdditionalRequest()) return null;
            MhtmlRemoteResourceResponse? response = await remoteResourcePolicy.ResourceFetcher!(
                new MhtmlRemoteResourceRequest(current, request.Source, request.Kind, redirectNumber),
                cancellationToken).ConfigureAwait(false);
            if (response == null) return null;
            if (response.RedirectLocation == null) {
                byte[]? bytes = response.EncodedBytes;
                return bytes == null
                    ? null
                    : new HtmlResolvedResource(bytes, response.ContentType, current, redirectNumber);
            }
            if (redirectNumber >= remoteResourcePolicy.MaximumRedirects) return null;
            current = response.RedirectLocation.IsAbsoluteUri
                ? response.RedirectLocation
                : new Uri(current, response.RedirectLocation);
        }
    }

    /// <summary>Serializes the archive to deterministic MHTML bytes.</summary>
    public byte[] ToBytes(EmailWriterOptions? options = null) =>
        new EmailDocumentWriter(options ?? EmailWriterOptions.Default).ToBytes(_mimeDocument, EmailFileFormat.Eml);

    /// <summary>Saves the archive to a file.</summary>
    public void Save(string path, EmailWriterOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        _mimeDocument.Save(path, EmailFileFormat.Eml, options);
    }

    /// <summary>Saves the archive to a caller-owned stream.</summary>
    public void Save(Stream stream, EmailWriterOptions? options = null) =>
        _mimeDocument.Save(stream, EmailFileFormat.Eml, options);

    /// <summary>Asynchronously saves the archive to a file.</summary>
    public Task<EmailWriteResult> SaveAsync(string path, EmailWriterOptions? options = null,
        CancellationToken cancellationToken = default) =>
        _mimeDocument.SaveAsync(path, EmailFileFormat.Eml, options, cancellationToken);

    /// <summary>Asynchronously saves the archive to a caller-owned stream.</summary>
    public Task<EmailWriteResult> SaveAsync(Stream stream, EmailWriterOptions? options = null,
        CancellationToken cancellationToken = default) =>
        _mimeDocument.SaveAsync(stream, EmailFileFormat.Eml, options, cancellationToken);

    private Task<HtmlResolvedResource?> ResolveResourceAsync(HtmlRenderResourceRequest request,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        MhtmlResource? resource = FindResource(request);
        return Task.FromResult(resource == null
            ? null
            : new HtmlResolvedResource(resource.EncodedContent, resource.ContentType));
    }

    private MhtmlResource? FindResource(HtmlRenderResourceRequest request) {
        string source = request.Source.Trim();
        string absolute = request.Uri.AbsoluteUri;
        if (request.Uri.Scheme.Equals("cid", StringComparison.OrdinalIgnoreCase)) {
            string contentId = Uri.UnescapeDataString(request.Uri.OriginalString.Substring("cid:".Length))
                .Trim().Trim('<', '>');
            return _resources.FirstOrDefault(resource => string.Equals(resource.ContentId, contentId,
                StringComparison.OrdinalIgnoreCase));
        }

        foreach (MhtmlResource resource in _resources) {
            if (!string.IsNullOrWhiteSpace(resource.ContentLocation)) {
                if (string.Equals(resource.ContentLocation, source, StringComparison.OrdinalIgnoreCase)) return resource;
                if (Uri.TryCreate(BaseUri, resource.ContentLocation, out Uri? resolved) &&
                    string.Equals(resolved.AbsoluteUri, absolute, StringComparison.OrdinalIgnoreCase)) return resource;
            }
            if (!string.IsNullOrWhiteSpace(resource.FileName) &&
                string.Equals(resource.FileName, source, StringComparison.OrdinalIgnoreCase)) return resource;
        }
        return null;
    }

    private static EmailDocument CreateMimeDocument(string html, IEnumerable<MhtmlResource> resources,
        string? contentLocation, string? rootContentId, string? subject) {
        var document = new EmailDocument {
            Format = EmailFileFormat.Eml,
            OutlookItemKind = OutlookItemKind.Message,
            Subject = subject
        };
        document.Body.Html = html;
        document.Body.HtmlCharset = "utf-8";
        document.Body.HtmlContentId = rootContentId;
        document.Body.HtmlContentLocation = contentLocation;
        document.Body.IsHtmlRelatedRoot = true;
        if (!string.IsNullOrWhiteSpace(contentLocation)) {
            document.Headers.Add(new EmailHeader("Snapshot-Content-Location", contentLocation!));
        }
        foreach (MhtmlResource resource in resources) document.Attachments.Add(resource.ToEmailAttachment());
        return document;
    }

    private static HtmlConversionDocumentOptions PrepareHtmlOptions(
        HtmlConversionDocumentOptions? source,
        Uri baseUri,
        IReadOnlyList<MhtmlResource> resources) {
        HtmlConversionDocumentOptions options = source?.Clone() ?? new HtmlConversionDocumentOptions();
        options.BaseUri ??= baseUri;
        HtmlUrlPolicy resourcePolicy = options.ResourceUrlPolicy.Clone();
        var archiveUris = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (MhtmlResource resource in resources) {
            if (!string.IsNullOrWhiteSpace(resource.ContentId)) {
                AddArchiveUri(archiveUris, "cid:" + resource.ContentId, baseUri);
            }
            AddArchiveUri(archiveUris, resource.ContentLocation, baseUri);
            AddArchiveUri(archiveUris, resource.FileName, baseUri);
        }

        if (resourcePolicy.RestrictUrlSchemes) {
            resourcePolicy.AllowedUrlSchemes.Add("cid");
            resourcePolicy.AllowedUrlSchemes.Add(baseUri.Scheme);
        }
        if (baseUri.IsFile) resourcePolicy.DisallowFileUrls = false;
        Func<string, string?>? callerTransform = resourcePolicy.ResolvedUrlTransform;
        resourcePolicy.ResolvedUrlTransform = resolved => {
            string? transformed = callerTransform == null ? resolved : callerTransform(resolved);
            if (string.IsNullOrWhiteSpace(transformed)
                || !Uri.TryCreate(transformed, UriKind.Absolute, out Uri? uri)) {
                return transformed;
            }

            bool archiveOnlyScheme = uri.Scheme.Equals("cid", StringComparison.OrdinalIgnoreCase)
                || baseUri.IsFile && uri.IsFile
                || (!baseUri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                    && !baseUri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                    && !baseUri.Scheme.Equals("data", StringComparison.OrdinalIgnoreCase)
                    && uri.Scheme.Equals(baseUri.Scheme, StringComparison.OrdinalIgnoreCase));
            return !archiveOnlyScheme || archiveUris.Contains(uri.AbsoluteUri) ? transformed : null;
        };
        options.ResourceUrlPolicy = resourcePolicy;
        return options;
    }

    private static void AddArchiveUri(HashSet<string> archiveUris, string? value, Uri baseUri) {
        if (string.IsNullOrWhiteSpace(value)) return;
        if (Uri.TryCreate(baseUri, value, out Uri? resolved)) archiveUris.Add(resolved.AbsoluteUri);
    }

    private static IReadOnlyList<EmailDiagnostic> BuildResourceDiagnostics(
        IReadOnlyList<MhtmlResource> resources,
        Uri baseUri,
        string? rootContentId,
        string? rootContentLocation) {
        var diagnostics = new List<EmailDiagnostic>();
        var contentIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var contentLocations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(rootContentId)) contentIds.Add(rootContentId!);
        if (!string.IsNullOrWhiteSpace(rootContentLocation)
            && Uri.TryCreate(baseUri, rootContentLocation, out _)) {
            contentLocations.Add(baseUri.AbsoluteUri);
        }
        for (int index = 0; index < resources.Count; index++) {
            MhtmlResource resource = resources[index];
            if (!string.IsNullOrWhiteSpace(resource.ContentId) && !contentIds.Add(resource.ContentId!)) {
                diagnostics.Add(new EmailDiagnostic(
                    MhtmlDiagnosticCodes.DuplicateContentId,
                    "Duplicate Content-ID was retained in archive order; the first resource is used for resolution.",
                    location: "resource[" + index + "]"));
            }
            if (string.IsNullOrWhiteSpace(resource.ContentLocation)) continue;
            if (!Uri.TryCreate(baseUri, resource.ContentLocation, out Uri? resolved)) {
                diagnostics.Add(new EmailDiagnostic(
                    MhtmlDiagnosticCodes.InvalidContentLocation,
                    "Content-Location could not be resolved against the archive base URI.",
                    location: "resource[" + index + "]"));
                continue;
            }
            if (!contentLocations.Add(resolved.AbsoluteUri)) {
                diagnostics.Add(new EmailDiagnostic(
                    MhtmlDiagnosticCodes.DuplicateContentLocation,
                    "Duplicate Content-Location was retained in archive order; the first resource is used for resolution.",
                    location: "resource[" + index + "]"));
            }
        }
        return diagnostics;
    }

    private static Uri ResolveBaseUri(string? contentLocation, Uri? sourceBaseUri) {
        if (!string.IsNullOrWhiteSpace(contentLocation)) {
            if (Uri.TryCreate(contentLocation, UriKind.Absolute, out Uri? absolute)) return absolute;
            Uri relativeBase = sourceBaseUri ?? FallbackBaseUri;
            if (Uri.TryCreate(relativeBase, contentLocation, out Uri? resolved)) return resolved;
        }
        return sourceBaseUri ?? FallbackBaseUri;
    }

    private static Uri? CreateFileBaseUri(string path) {
        try {
            return new Uri(Path.GetFullPath(path));
        } catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException ||
                                           exception is PathTooLongException) {
            return null;
        }
    }

    private static bool IsMultipartRelated(IEnumerable<EmailHeader> headers) {
        string? contentType = GetHeaderValue(headers, "Content-Type");
        return contentType != null && contentType.TrimStart()
            .StartsWith("multipart/related", StringComparison.OrdinalIgnoreCase);
    }

    private static string? GetHeaderValue(IEnumerable<EmailHeader> headers, string name) =>
        headers.FirstOrDefault(header => string.Equals(header.Name, name, StringComparison.OrdinalIgnoreCase))?.Value;

    private static InvalidDataException CreateReadException(IEnumerable<EmailDiagnostic> diagnostics) {
        EmailDiagnostic? error = diagnostics.FirstOrDefault(diagnostic =>
            diagnostic.Severity == EmailDiagnosticSeverity.Error);
        return error == null
            ? new InvalidDataException("The MHTML archive could not be read.")
            : new InvalidDataException(string.Concat("The MHTML archive could not be read: ", error.Code,
                ": ", error.Message));
    }

    private static string? NormalizeOptional(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value!.Trim();

    private static string? NormalizeContentId(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value!.Trim().Trim('<', '>');

    private sealed class ManagedResourceResolver {
        private readonly HtmlUrlPolicy _fallbackResourceUrlPolicy;
        private readonly MhtmlRemoteResourcePolicy _remoteResourcePolicy;
        private readonly bool _includeEmbeddedResources;

        internal ManagedResourceResolver(
            MhtmlDocument document,
            HtmlUrlPolicy fallbackResourceUrlPolicy,
            MhtmlRemoteResourcePolicy remoteResourcePolicy,
            bool includeEmbeddedResources) {
            Document = document;
            _fallbackResourceUrlPolicy = fallbackResourceUrlPolicy;
            _remoteResourcePolicy = remoteResourcePolicy;
            _includeEmbeddedResources = includeEmbeddedResources;
        }

        internal MhtmlDocument Document { get; }

        internal ManagedResourceResolver WithEmbeddedResolution(bool includeEmbeddedResources) =>
            includeEmbeddedResources == _includeEmbeddedResources
                ? this
                : new ManagedResourceResolver(
                    Document,
                    _fallbackResourceUrlPolicy,
                    _remoteResourcePolicy,
                    includeEmbeddedResources);

        internal async Task<HtmlResolvedResource?> ResolveAsync(
            HtmlRenderResourceRequest request,
            CancellationToken cancellationToken) {
            HtmlResolvedResource? embedded = _includeEmbeddedResources
                ? await Document.ResolveResourceAsync(request, cancellationToken).ConfigureAwait(false)
                : null;
            if (embedded != null || _remoteResourcePolicy.ResourceFetcher == null) return embedded;
            return await Document.ResolveRemoteResourceAsync(
                request,
                _fallbackResourceUrlPolicy,
                _remoteResourcePolicy,
                cancellationToken).ConfigureAwait(false);
        }
    }
}
