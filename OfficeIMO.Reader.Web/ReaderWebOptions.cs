namespace OfficeIMO.Reader.Web;

/// <summary>Controls bounded HTTP retrieval before content enters an existing Reader pipeline.</summary>
public sealed class ReaderWebOptions {
    /// <summary>Default maximum downloaded response size: 64 MiB.</summary>
    public const long DefaultMaxResponseBytes = 64L * 1024L * 1024L;

    /// <summary>Maximum supported response bound: 1 GiB.</summary>
    public const long MaximumResponseBytes = 1024L * 1024L * 1024L;

    /// <summary>Gets or sets the maximum downloaded response bytes. Default: 64 MiB.</summary>
    public long MaxResponseBytes { get; set; } = DefaultMaxResponseBytes;

    /// <summary>Gets or sets the timeout covering response headers and body download. Default: two minutes.</summary>
    public TimeSpan RequestTimeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>
    /// Gets or sets the maximum concurrent web read operations on one
    /// <see cref="OfficeDocumentWebReader"/>. Default: 4.
    /// </summary>
    public int MaxConcurrentRequests { get; set; } = 4;

    /// <summary>
    /// Gets or sets an optional exact host allowlist. An empty list permits any host that passes the
    /// scheme and non-public IP-literal checks.
    /// </summary>
    public IReadOnlyList<string> AllowedHosts { get; set; } = Array.Empty<string>();

    /// <summary>Gets or sets whether subdomains of entries in <see cref="AllowedHosts"/> are permitted.</summary>
    public bool AllowSubdomains { get; set; }

    /// <summary>
    /// Gets or sets whether localhost names and loopback, private, link-local, or non-routable IP literals are permitted.
    /// Default: false.
    /// </summary>
    /// <remarks>
    /// This option screens URI literals only. It does not resolve hostnames or intercept redirects.
    /// When a URI is not trusted, the caller-provided HTTP handler must validate resolved addresses at
    /// connection time and validate every redirect destination before sending the redirected request.
    /// </remarks>
    public bool AllowLocalhostAndNonPublicIpLiterals { get; set; }

    /// <summary>
    /// Gets or sets whether query strings are retained in transport metadata. Default: false.
    /// Keep disabled when URLs may contain signed tokens or other secrets.
    /// </summary>
    public bool IncludeQueryInMetadata { get; set; }

    internal ReaderWebOptions CloneValidated() {
        if (MaxResponseBytes < 1 || MaxResponseBytes > MaximumResponseBytes) {
            throw new ArgumentOutOfRangeException(nameof(MaxResponseBytes));
        }
        if (RequestTimeout <= TimeSpan.Zero || RequestTimeout > TimeSpan.FromDays(1)) {
            throw new ArgumentOutOfRangeException(nameof(RequestTimeout));
        }
        if (MaxConcurrentRequests < 1 || MaxConcurrentRequests > 64) {
            throw new ArgumentOutOfRangeException(nameof(MaxConcurrentRequests));
        }
        if (AllowedHosts == null) throw new ArgumentNullException(nameof(AllowedHosts));

        string[] allowedHosts = AllowedHosts
            .Select(NormalizeAllowedHost)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        return new ReaderWebOptions {
            MaxResponseBytes = MaxResponseBytes,
            RequestTimeout = RequestTimeout,
            MaxConcurrentRequests = MaxConcurrentRequests,
            AllowedHosts = allowedHosts,
            AllowSubdomains = AllowSubdomains,
            AllowLocalhostAndNonPublicIpLiterals = AllowLocalhostAndNonPublicIpLiterals,
            IncludeQueryInMetadata = IncludeQueryInMetadata
        };
    }

    private static string NormalizeAllowedHost(string host) {
        if (string.IsNullOrWhiteSpace(host)) {
            throw new ArgumentException("Allowed host values cannot be empty.", nameof(AllowedHosts));
        }
        string normalized = host.Trim().TrimEnd('.');
        if (normalized.Length > 253 ||
            normalized.IndexOf("://", StringComparison.Ordinal) >= 0 ||
            normalized.IndexOfAny(new[] { '/', '\\', '*', '?', '#', '@' }) >= 0) {
            throw new ArgumentException("Allowed host values must be host names or IP literals without schemes or wildcards.", nameof(AllowedHosts));
        }
        if (normalized.Length > 1 && normalized[0] == '[' && normalized[normalized.Length - 1] == ']') {
            normalized = normalized.Substring(1, normalized.Length - 2);
        }
        if (IPAddress.TryParse(normalized, out IPAddress? address)) {
            return address.ToString().ToLowerInvariant();
        }
        if (Uri.CheckHostName(normalized) == UriHostNameType.Unknown) {
            throw new ArgumentException("Allowed host value is not a valid host name or IP literal: " + host, nameof(AllowedHosts));
        }
        try {
            return new Uri(Uri.UriSchemeHttp + "://" + normalized).IdnHost.TrimEnd('.').ToLowerInvariant();
        } catch (UriFormatException exception) {
            throw new ArgumentException("Allowed host value is not a valid host name: " + host, nameof(AllowedHosts), exception);
        }
    }
}
