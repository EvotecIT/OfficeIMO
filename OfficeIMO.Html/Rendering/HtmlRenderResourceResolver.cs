namespace OfficeIMO.Html;

/// <summary>
/// Asynchronously resolves a policy-approved HTML render resource without prescribing a network or storage dependency.
/// </summary>
/// <param name="request">Resolved resource request.</param>
/// <param name="cancellationToken">Cancellation token covering caller cancellation and the configured resource timeout.</param>
/// <returns>Resolved bytes and media type, or null when the caller cannot provide the resource.</returns>
public delegate Task<HtmlResolvedResource?> HtmlRenderResourceResolver(HtmlRenderResourceRequest request, CancellationToken cancellationToken);

internal delegate bool HtmlRenderSynchronousResourceResolver(
    HtmlRenderResourceRequest request,
    CancellationToken cancellationToken,
    out HtmlResolvedResource? resource);

internal sealed class HtmlRenderResourceByteLimitException : Exception {
    internal HtmlRenderResourceByteLimitException(long actualBytes) {
        ActualBytes = actualBytes;
    }

    internal long ActualBytes { get; }
}

/// <summary>
/// Policy-approved resource request passed to an application-supplied resolver.
/// </summary>
public sealed class HtmlRenderResourceRequest {
    internal HtmlRenderResourceRequest(Uri uri, string source, HtmlResourceKind kind) {
        Uri = uri;
        Source = source;
        Kind = kind;
    }

    /// <summary>Absolute URI after base-URI resolution and URL-policy evaluation.</summary>
    public Uri Uri { get; }

    /// <summary>Raw source reference from the HTML document.</summary>
    public string Source { get; }

    /// <summary>Resource kind requested by the renderer.</summary>
    public HtmlResourceKind Kind { get; }
}

/// <summary>
/// Immutable bytes returned by an application-supplied HTML resource resolver.
/// </summary>
public sealed class HtmlResolvedResource {
    private readonly byte[] _bytes;

    /// <summary>Creates a resolved resource snapshot.</summary>
    public HtmlResolvedResource(byte[] bytes, string contentType) {
        if (bytes == null || bytes.Length == 0) throw new ArgumentException("Resolved resources require non-empty bytes.", nameof(bytes));
        _bytes = (byte[])bytes.Clone();
        ContentType = string.IsNullOrWhiteSpace(contentType) ? "application/octet-stream" : contentType.Trim();
    }

    /// <summary>Resolved resource bytes.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Declared media type.</summary>
    public string ContentType { get; }
}
