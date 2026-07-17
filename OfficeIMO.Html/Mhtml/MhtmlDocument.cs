using OfficeIMO.Email;

namespace OfficeIMO.Html;

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
        _mimeDiagnostics = Array.Empty<EmailDiagnostic>();
        ContentLocation = NormalizeOptional(contentLocation);
        RootContentId = NormalizeContentId(rootContentId);
        Subject = NormalizeOptional(subject);
        BaseUri = ResolveBaseUri(ContentLocation, null);
        HtmlDocument = HtmlConversionDocument.Parse(html, PrepareHtmlOptions(htmlOptions, BaseUri));
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
        _resources = _mimeDocument.Attachments.Where(static attachment => attachment.IsMimeRelated)
            .Select(MhtmlResource.FromEmailAttachment)
            .ToArray();
        _mimeDiagnostics = readResult.Diagnostics;
        HtmlDocument = HtmlConversionDocument.Parse(html, PrepareHtmlOptions(htmlOptions, BaseUri));
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
    /// The hyperlink policy is left unchanged, and an existing resolver remains the fallback for resources absent
    /// from the archive.
    /// </summary>
    public void ConfigureRenderOptions(HtmlRenderOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.BaseUri ??= BaseUri;
        options.UrlPolicy ??= HtmlUrlPolicy.CreateOfficeIMOProfile();
        HtmlUrlPolicy resourceUrlPolicy = (options.ResourceUrlPolicy ?? options.UrlPolicy).Clone();
        if (resourceUrlPolicy.RestrictUrlSchemes) {
            resourceUrlPolicy.AllowedUrlSchemes.Add("cid");
            resourceUrlPolicy.AllowedUrlSchemes.Add(BaseUri.Scheme);
        }
        resourceUrlPolicy.DisallowFileUrls = false;
        options.ResourceUrlPolicy = resourceUrlPolicy;
        HtmlRenderResourceResolver embeddedResolver = CreateResourceResolver();
        HtmlRenderResourceResolver? fallbackResolver = options.ResourceResolver;
        options.ResourceResolver = async (request, cancellationToken) => {
            HtmlResolvedResource? embedded = await embeddedResolver(request, cancellationToken).ConfigureAwait(false);
            if (embedded != null || fallbackResolver == null) return embedded;
            return await fallbackResolver(request, cancellationToken).ConfigureAwait(false);
        };
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
            : new HtmlResolvedResource(resource.Content, resource.ContentType));
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

    private static HtmlConversionDocumentOptions PrepareHtmlOptions(HtmlConversionDocumentOptions? source,
        Uri baseUri) {
        HtmlConversionDocumentOptions options = source?.Clone() ?? new HtmlConversionDocumentOptions();
        options.BaseUri ??= baseUri;
        return options;
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
}
