using OfficeIMO.Html;

namespace OfficeIMO.Mhtml;

/// <summary>
/// Fetches exactly one remote response without automatically following redirects. MHTML owns
/// redirect traversal so policy can approve every destination before the next request is made.
/// </summary>
public delegate Task<MhtmlRemoteResourceResponse?> MhtmlRemoteResourceFetcher(
    MhtmlRemoteResourceRequest request,
    CancellationToken cancellationToken);

/// <summary>A single policy-approved remote-resource request.</summary>
public sealed class MhtmlRemoteResourceRequest {
    internal MhtmlRemoteResourceRequest(Uri uri, string source, HtmlResourceKind kind, int redirectNumber) {
        Uri = uri;
        Source = source;
        Kind = kind;
        RedirectNumber = redirectNumber;
    }

    /// <summary>Absolute URI approved for this one network request.</summary>
    public Uri Uri { get; }

    /// <summary>Original source reference from the HTML document.</summary>
    public string Source { get; }

    /// <summary>Requested resource kind.</summary>
    public HtmlResourceKind Kind { get; }

    /// <summary>Zero for the initial request, then one-based for each approved redirect hop.</summary>
    public int RedirectNumber { get; }
}

/// <summary>One remote response containing either bytes or a redirect location.</summary>
public sealed class MhtmlRemoteResourceResponse {
    private readonly byte[]? _bytes;

    /// <summary>Creates a successful resource response.</summary>
    public MhtmlRemoteResourceResponse(byte[] bytes, string contentType) {
        if (bytes == null || bytes.Length == 0) {
            throw new ArgumentException("Remote resource responses require non-empty bytes.", nameof(bytes));
        }
        _bytes = (byte[])bytes.Clone();
        ContentType = string.IsNullOrWhiteSpace(contentType) ? "application/octet-stream" : contentType.Trim();
    }

    private MhtmlRemoteResourceResponse(Uri redirectLocation) {
        RedirectLocation = redirectLocation ?? throw new ArgumentNullException(nameof(redirectLocation));
        ContentType = "application/octet-stream";
    }

    /// <summary>Creates a redirect response. Relative locations are resolved against the requested URI.</summary>
    public static MhtmlRemoteResourceResponse Redirect(Uri location) => new MhtmlRemoteResourceResponse(location);

    /// <summary>Response bytes, or null for a redirect.</summary>
    public byte[]? Bytes => _bytes == null ? null : (byte[])_bytes.Clone();

    internal byte[]? EncodedBytes => _bytes;

    /// <summary>Declared media type for a successful response.</summary>
    public string ContentType { get; }

    /// <summary>Redirect location, or null for a successful response.</summary>
    public Uri? RedirectLocation { get; }
}
