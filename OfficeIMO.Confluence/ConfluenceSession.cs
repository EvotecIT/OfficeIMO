namespace OfficeIMO.Confluence;

/// <summary>Options shared by Confluence Cloud requests.</summary>
public sealed class ConfluenceSessionOptions {
    /// <summary>Confluence site URL, for example <c>https://example.atlassian.net/wiki/</c>. Direct API requests use its origin.</summary>
    public Uri SiteUri { get; set; } = null!;
    /// <summary>
    /// Atlassian Cloud identifier used by OAuth 2.0 (3LO). When supplied, requests are routed through
    /// <c>https://api.atlassian.com/ex/confluence/{cloudId}/</c> while <see cref="SiteUri"/> remains the human-facing site.
    /// </summary>
    public string? CloudId { get; set; }
    public string ApplicationName { get; set; } = "OfficeIMO";
    public HttpClient? HttpClient { get; set; }
    public TimeSpan RequestTimeout { get; set; } = TimeSpan.FromSeconds(100);
    public int MaxRetryCount { get; set; } = 3;
    public TimeSpan RetryBaseDelay { get; set; } = TimeSpan.FromMilliseconds(250);
    public TimeSpan RetryMaxDelay { get; set; } = TimeSpan.FromSeconds(8);
}

/// <summary>A configured Confluence Cloud session.</summary>
public sealed class ConfluenceSession {
    private readonly ConfluenceSessionOptions _options;

    public ConfluenceSession(IConfluenceCredentialSource credentialSource, ConfluenceSessionOptions options) {
        CredentialSource = credentialSource ?? throw new ArgumentNullException(nameof(credentialSource));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (options.SiteUri == null || !options.SiteUri.IsAbsoluteUri) throw new ArgumentException("An absolute Confluence site URI is required.", nameof(options));
        if (!string.Equals(options.SiteUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Confluence Cloud site URI must use HTTPS.", nameof(options));
        if (!string.IsNullOrEmpty(options.SiteUri.UserInfo)) throw new ArgumentException("Confluence site URI cannot contain embedded credentials.", nameof(options));
        if (options.CloudId != null && string.IsNullOrWhiteSpace(options.CloudId)) throw new ArgumentException("Confluence Cloud ID cannot be empty.", nameof(options));
        if (options.MaxRetryCount < 0) throw new ArgumentOutOfRangeException(nameof(options), "Max retry count cannot be negative.");
        if (options.RequestTimeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(options), "Request timeout must be greater than zero.");
        if (options.RetryBaseDelay < TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(options), "Retry base delay cannot be negative.");
        if (options.RetryMaxDelay < TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(options), "Retry maximum delay cannot be negative.");

        _options = Clone(options);
        ApiBaseUri = string.IsNullOrWhiteSpace(_options.CloudId)
            ? GetOriginRoot(_options.SiteUri)
            : new Uri("https://api.atlassian.com/ex/confluence/" + Uri.EscapeDataString(_options.CloudId!) + "/", UriKind.Absolute);
    }

    public IConfluenceCredentialSource CredentialSource { get; }
    /// <summary>Returns a defensive copy of the validated session options.</summary>
    public ConfluenceSessionOptions Options => Clone(_options);
    /// <summary>Resolved API base used for direct-site or OAuth cloud-ID requests.</summary>
    public Uri ApiBaseUri { get; }
    internal ConfluenceSessionOptions RuntimeOptions => _options;

    /// <summary>Creates a client for this session.</summary>
    public ConfluenceClient CreateClient() => new ConfluenceClient(this);

    private static ConfluenceSessionOptions Clone(ConfluenceSessionOptions source) => new ConfluenceSessionOptions {
        SiteUri = source.SiteUri,
        CloudId = source.CloudId?.Trim(),
        ApplicationName = source.ApplicationName,
        HttpClient = source.HttpClient,
        RequestTimeout = source.RequestTimeout,
        MaxRetryCount = source.MaxRetryCount,
        RetryBaseDelay = source.RetryBaseDelay,
        RetryMaxDelay = source.RetryMaxDelay,
    };

    private static Uri GetOriginRoot(Uri uri) => new Uri(uri.GetLeftPart(UriPartial.Authority) + "/", UriKind.Absolute);
}
