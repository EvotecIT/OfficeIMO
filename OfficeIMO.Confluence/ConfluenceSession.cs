namespace OfficeIMO.Confluence;

/// <summary>Options shared by Confluence Cloud requests.</summary>
public sealed class ConfluenceSessionOptions {
    /// <summary>Confluence site root, for example <c>https://example.atlassian.net/</c>.</summary>
    public Uri SiteUri { get; set; } = null!;
    public string ApplicationName { get; set; } = "OfficeIMO";
    public HttpClient? HttpClient { get; set; }
    public TimeSpan RequestTimeout { get; set; } = TimeSpan.FromSeconds(100);
    public int MaxRetryCount { get; set; } = 3;
    public TimeSpan RetryBaseDelay { get; set; } = TimeSpan.FromMilliseconds(250);
    public TimeSpan RetryMaxDelay { get; set; } = TimeSpan.FromSeconds(8);
}

/// <summary>A configured Confluence Cloud session.</summary>
public sealed class ConfluenceSession {
    public ConfluenceSession(IConfluenceCredentialSource credentialSource, ConfluenceSessionOptions options) {
        CredentialSource = credentialSource ?? throw new ArgumentNullException(nameof(credentialSource));
        Options = options ?? throw new ArgumentNullException(nameof(options));
        if (options.SiteUri == null || !options.SiteUri.IsAbsoluteUri) throw new ArgumentException("An absolute Confluence site URI is required.", nameof(options));
        if (!string.Equals(options.SiteUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Confluence Cloud site URI must use HTTPS.", nameof(options));
        if (!string.IsNullOrEmpty(options.SiteUri.UserInfo)) throw new ArgumentException("Confluence site URI cannot contain embedded credentials.", nameof(options));
        if (options.MaxRetryCount < 0) throw new ArgumentOutOfRangeException(nameof(options), "Max retry count cannot be negative.");
        if (options.RequestTimeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(options), "Request timeout must be greater than zero.");
        if (options.RetryBaseDelay < TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(options), "Retry base delay cannot be negative.");
        if (options.RetryMaxDelay < TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(options), "Retry maximum delay cannot be negative.");
    }

    public IConfluenceCredentialSource CredentialSource { get; }
    public ConfluenceSessionOptions Options { get; }

    /// <summary>Creates a client for this session.</summary>
    public ConfluenceClient CreateClient() => new ConfluenceClient(this);
}
