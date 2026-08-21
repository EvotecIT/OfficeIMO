using OfficeIMO.Html;

namespace OfficeIMO.Mhtml;

/// <summary>
/// Explicit policy for resources absent from an MHTML archive and delegated to a caller resolver.
/// Embedded MIME resources are not network resources and remain governed by the shared HTML byte/count limits.
/// </summary>
public sealed class MhtmlRemoteResourcePolicy {
    private readonly HashSet<string> _allowedOrigins = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    /// <summary>Creates the default offline profile. No remote fallback resolver is invoked.</summary>
    public static MhtmlRemoteResourcePolicy CreateEmbeddedOnlyProfile() => new MhtmlRemoteResourcePolicy();

    /// <summary>
    /// Creates a bounded same-origin remote profile. Additional origins can be added through
    /// <see cref="AllowedOrigins"/> using absolute HTTP or HTTPS origins.
    /// </summary>
    public static MhtmlRemoteResourcePolicy CreateSameOriginProfile(int maximumRedirects = 3) =>
        new MhtmlRemoteResourcePolicy {
            AllowRemoteResources = true,
            AllowSameOrigin = true,
            MaximumRedirects = maximumRedirects
        };

    /// <summary>Whether a caller-supplied resolver may be invoked for HTTP or HTTPS resources.</summary>
    public bool AllowRemoteResources { get; set; }

    /// <summary>Whether the archive root origin is allowed. Defaults to false in the offline profile.</summary>
    public bool AllowSameOrigin { get; set; }

    /// <summary>Maximum redirects a resolver may follow. Defaults to zero.</summary>
    public int MaximumRedirects { get; set; }

    /// <summary>
    /// Optional one-hop remote fetcher. It must return redirects without following them; OfficeIMO
    /// validates each next URI before invoking the fetcher again.
    /// </summary>
    public MhtmlRemoteResourceFetcher? ResourceFetcher { get; set; }

    /// <summary>Additional allowed origins, expressed as absolute HTTP or HTTPS URI strings.</summary>
    public ISet<string> AllowedOrigins => _allowedOrigins;

    /// <summary>Maximum bytes accepted from one resource while remote fallback is enabled.</summary>
    public long MaximumResourceBytes { get; set; } = 10L * 1024L * 1024L;

    /// <summary>Maximum total resource bytes accepted during one render while remote fallback is enabled.</summary>
    public long MaximumTotalResourceBytes { get; set; } = 50L * 1024L * 1024L;

    /// <summary>Maximum resources accepted during one render while remote fallback is enabled.</summary>
    public int MaximumResourceCount { get; set; } = 256;

    /// <summary>Maximum resolver invocations attempted during one render while remote fallback is enabled.</summary>
    public int MaximumResourceRequests { get; set; } = 512;

    /// <summary>Maximum duration of one resolver invocation while remote fallback is enabled.</summary>
    public TimeSpan ResourceTimeout { get; set; } = TimeSpan.FromSeconds(30D);

    internal void ApplyLimits(HtmlRenderOptions options) {
        Validate();
        if (!AllowRemoteResources || ResourceFetcher == null) return;
        options.MaxResourceBytes = Math.Min(options.MaxResourceBytes, MaximumResourceBytes);
        options.MaxTotalResourceBytes = Math.Min(options.MaxTotalResourceBytes, MaximumTotalResourceBytes);
        options.MaxResourceCount = Math.Min(options.MaxResourceCount, MaximumResourceCount);
        options.MaxResourceRequests = Math.Min(options.MaxResourceRequests, MaximumResourceRequests);
        if (options.ResourceTimeout > ResourceTimeout) options.ResourceTimeout = ResourceTimeout;
    }

    internal bool AllowsRequest(Uri uri, Uri archiveBaseUri) {
        if (!AllowRemoteResources || !IsRemote(uri)) return false;
        return IsAllowedOrigin(uri, archiveBaseUri);
    }

    private bool IsAllowedOrigin(Uri uri, Uri archiveBaseUri) {
        string origin = GetOrigin(uri);
        if (AllowSameOrigin && IsRemote(archiveBaseUri) &&
            string.Equals(origin, GetOrigin(archiveBaseUri), StringComparison.OrdinalIgnoreCase)) return true;
        return _allowedOrigins.Any(value => TryNormalizeOrigin(value, out string? allowed) &&
            string.Equals(origin, allowed, StringComparison.OrdinalIgnoreCase));
    }

    private void Validate() {
        if (MaximumRedirects < 0) throw new ArgumentOutOfRangeException(nameof(MaximumRedirects));
        if (MaximumResourceBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaximumResourceBytes));
        if (MaximumTotalResourceBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaximumTotalResourceBytes));
        if (MaximumTotalResourceBytes < MaximumResourceBytes) {
            throw new ArgumentOutOfRangeException(nameof(MaximumTotalResourceBytes),
                "The total resource-byte limit must be at least the per-resource byte limit.");
        }
        if (MaximumResourceCount < 1) throw new ArgumentOutOfRangeException(nameof(MaximumResourceCount));
        if (MaximumResourceRequests < 1) throw new ArgumentOutOfRangeException(nameof(MaximumResourceRequests));
        if (MaximumResourceRequests < MaximumResourceCount) {
            throw new ArgumentOutOfRangeException(nameof(MaximumResourceRequests),
                "The resource-request limit must be at least the accepted-resource count limit.");
        }
        if (ResourceTimeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(ResourceTimeout));
        foreach (string origin in _allowedOrigins) {
            if (!TryNormalizeOrigin(origin, out _)) {
                throw new ArgumentException("AllowedOrigins entries must be absolute HTTP or HTTPS origins.", nameof(AllowedOrigins));
            }
        }
    }

    private static bool TryNormalizeOrigin(string? value, out string? origin) {
        origin = null;
        if (string.IsNullOrWhiteSpace(value) || !Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) || !IsRemote(uri)) return false;
        origin = GetOrigin(uri);
        return string.Equals(uri.AbsoluteUri.TrimEnd('/'), origin, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsRemote(Uri uri) => uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
        uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase);

    private static string GetOrigin(Uri uri) => uri.GetLeftPart(UriPartial.Authority).TrimEnd('/');
}
