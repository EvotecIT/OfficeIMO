namespace OfficeIMO.Html.Runtime;

/// <summary>Session-owned HTTP resource authority and cumulative loading limits.</summary>
public sealed class HtmlRuntimeResourcePolicy {
    /// <summary>Allows HTTP requests when no supplied resource matches. False keeps loading offline.</summary>
    public bool AllowNetwork { get; set; }
    /// <summary>Additional permitted HTTP(S) origins. The document origin is always permitted; this does not grant browser CORS permission.</summary>
    public IReadOnlyList<Uri> AllowedOrigins { get; set; } = Array.Empty<Uri>();
    /// <summary>Deadline for each load, including queue admission, redirects and response reading.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(5);
    /// <summary>Maximum simultaneous loads.</summary>
    public int MaxConcurrentRequests { get; set; } = 4;
    /// <summary>Maximum attempted requests over the session, including supplied resources and redirect hops.</summary>
    public int MaxRequests { get; set; } = 128;
    /// <summary>Maximum encoded bytes in one supplied or loaded resource.</summary>
    public long MaxResourceBytes { get; set; } = 4 * 1024 * 1024;
    /// <summary>Maximum combined supplied resource bytes, and separately the cumulative bytes read by loads over the session.</summary>
    public long MaxTotalBytes { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum encoded fetch request body size.</summary>
    public long MaxRequestBytes { get; set; } = 1024 * 1024;
    /// <summary>Maximum request body bytes sent over the session, including redirect replays.</summary>
    public long MaxTotalRequestBytes { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum redirects in one load. Every target is checked before requesting it.</summary>
    public int MaxRedirects { get; set; } = 5;

    internal HtmlRuntimeResourcePolicy Snapshot() {
        if (Timeout <= TimeSpan.Zero || Timeout > TimeSpan.FromMinutes(5) || MaxConcurrentRequests <= 0 || MaxRequests <= 0 ||
            MaxResourceBytes <= 0 || MaxResourceBytes > int.MaxValue || MaxTotalBytes <= 0 || MaxRedirects < 0 || MaxRequestBytes <= 0 || MaxRequestBytes > int.MaxValue || MaxTotalRequestBytes <= 0)
            throw new ArgumentException("Resource limits must be positive and within supported ranges.");
        ArgumentNullException.ThrowIfNull(AllowedOrigins);
        if (AllowedOrigins.Count > MaxRequests) throw new ArgumentException("Too many resource origins.");
        var origins = AllowedOrigins.Select(url => {
            ValidateUrl(url);
            if (url.AbsolutePath != "/" || url.Query.Length != 0 || url.Fragment.Length != 0) throw new ArgumentException("AllowedOrigins must contain origins without paths, queries or fragments.");
            return new Uri(url.GetLeftPart(UriPartial.Authority));
        }).ToArray();
        return new HtmlRuntimeResourcePolicy { AllowNetwork = AllowNetwork, AllowedOrigins = origins, Timeout = Timeout,
            MaxConcurrentRequests = MaxConcurrentRequests, MaxRequests = MaxRequests, MaxResourceBytes = MaxResourceBytes,
            MaxTotalBytes = MaxTotalBytes, MaxRedirects = MaxRedirects, MaxRequestBytes = MaxRequestBytes, MaxTotalRequestBytes = MaxTotalRequestBytes };
    }

    internal static Uri ValidateUrl(Uri url) {
        ArgumentNullException.ThrowIfNull(url);
        if (!url.IsAbsoluteUri || (url.Scheme != Uri.UriSchemeHttp && url.Scheme != Uri.UriSchemeHttps) ||
            url.UserInfo.Length != 0 || url.AbsoluteUri.Length > 8192)
            throw new ArgumentException("Runtime resource URLs must be absolute HTTP(S) URLs without credentials.");
        return url;
    }

    internal static string Key(Uri url) => new UriBuilder(ValidateUrl(url)) { Fragment = string.Empty }.Uri.AbsoluteUri;
    internal static string Origin(Uri url) => ValidateUrl(url).GetLeftPart(UriPartial.Authority);
}
