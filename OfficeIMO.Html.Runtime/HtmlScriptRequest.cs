using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Runtime;

/// <summary>A trusted scripted session with a versioned behavior profile and explicit resource authority.</summary>
public sealed class HtmlScriptRequest {
    /// <summary>Versioned behavior contract enabled for this session.</summary>
    public HtmlRuntimeProfile Profile { get; set; } = HtmlRuntimeProfile.ScriptedDocumentV1;
    /// <summary>Complete HTML source. Inline classic scripts execute while the document loads.</summary>
    public string Html { get; set; } = string.Empty;
    /// <summary>Initial document identity and base for relative resource URLs.</summary>
    public Uri DocumentUrl { get; set; } = new("https://officeimo.invalid/");
    /// <summary>Optional immutable resources for offline scripts, stylesheets and other document assets.</summary>
    public IReadOnlyList<HtmlRuntimeResource> Resources { get; set; } = Array.Empty<HtmlRuntimeResource>();
    /// <summary>Network authority and cumulative resource budgets. Network access is disabled by default.</summary>
    public HtmlRuntimeResourcePolicy ResourcePolicy { get; set; } = new();
    /// <summary>Classic scripts executed in order after document loading.</summary>
    public IReadOnlyList<string> Scripts { get; set; } = Array.Empty<string>();
    /// <summary>A JavaScript expression which must evaluate to boolean true before capture.</summary>
    public string ReadyExpression { get; set; } = "true";
    /// <summary>Deadline per session command, including queue admission and capture transfer; total deadline for CaptureTrustedAsync.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(10);
    /// <summary>Maximum live worker lifetime, including idle time, from startup until termination.</summary>
    public TimeSpan SessionTimeout { get; set; } = TimeSpan.FromMinutes(5);
    /// <summary>Interval between readiness checks. This does not imply network or layout stability.</summary>
    public TimeSpan PollInterval { get; set; } = TimeSpan.FromMilliseconds(10);
    /// <summary>Combined UTF-16 source budget for HTML, supplied scripts and the readiness expression.</summary>
    public int MaxInputCharacters { get; set; } = 8 * 1024 * 1024;
    /// <summary>UTF-16 budget for the complete serialized worker response.</summary>
    public int MaxOutputCharacters { get; set; } = 8 * 1024 * 1024;
    /// <summary>Maximum attached nodes in the captured document, including template fragments.</summary>
    public int MaxNodes { get; set; } = 100_000;
    /// <summary>Maximum captured element nesting depth.</summary>
    public int MaxDepth { get; set; } = 256;
    /// <summary>Maximum tracked promise rejections awaiting a handler within one script turn.</summary>
    public int MaxPendingPromiseRejections { get; set; } = 1024;
    /// <summary>Maximum retained UTF-16 key and value characters in each session-local Web Storage area.</summary>
    public int MaxStorageCharacters { get; set; } = 1024 * 1024;
    /// <summary>Maximum distinct module sources retained per session, including inline roots and failed loads.</summary>
    public int MaxModuleCount { get; set; } = 1024;
    /// <summary>Maximum cross-document loads and reloads admitted over the session lifetime.</summary>
    public int MaxNavigations { get; set; } = 128;
    /// <summary>Maximum retained same-document history entries, including the first and current entries.</summary>
    public int MaxHistoryEntries { get; set; } = 128;
    /// <summary>Maximum estimated serialized bytes in one history state graph.</summary>
    public int MaxHistoryStateBytes { get; set; } = 1024 * 1024;
    /// <summary>Maximum estimated serialized bytes retained across all history state graphs.</summary>
    public int MaxHistoryTotalStateBytes { get; set; } = 8 * 1024 * 1024;
    /// <summary>Maximum queued history traversals and fragment-change notifications.</summary>
    public int MaxPendingHistoryTasks { get; set; } = 1024;
    /// <summary>Layout viewport width in CSS pixels for WebApplicationV1 inspection and actionability.</summary>
    public double ViewportWidth { get; set; } = 1280D;
    /// <summary>Layout viewport height in CSS pixels for WebApplicationV1 inspection and actionability.</summary>
    public double ViewportHeight { get; set; } = 720D;

    internal HtmlScriptRequest Snapshot() {
        if (Html == null || Scripts == null || ReadyExpression == null) throw new ArgumentException("HTML, scripts and readiness are required.");
        if (!Enum.IsDefined(Profile)) throw new ArgumentOutOfRangeException(nameof(Profile));
        if (Timeout <= TimeSpan.Zero || Timeout > TimeSpan.FromMinutes(5)) throw new ArgumentOutOfRangeException(nameof(Timeout));
        if (SessionTimeout <= TimeSpan.Zero || SessionTimeout > TimeSpan.FromHours(1)) throw new ArgumentOutOfRangeException(nameof(SessionTimeout));
        if (PollInterval < TimeSpan.FromMilliseconds(1) || PollInterval > Timeout) throw new ArgumentOutOfRangeException(nameof(PollInterval));
        if (MaxInputCharacters <= 0 || MaxOutputCharacters <= 0 || MaxNodes <= 0 || MaxDepth <= 0 || MaxPendingPromiseRejections <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputCharacters), "Resource limits must be positive.");
        if (MaxStorageCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxStorageCharacters));
        if (MaxModuleCount <= 0) throw new ArgumentOutOfRangeException(nameof(MaxModuleCount));
        if (MaxNavigations <= 0) throw new ArgumentOutOfRangeException(nameof(MaxNavigations));
        if (MaxHistoryEntries < 2) throw new ArgumentOutOfRangeException(nameof(MaxHistoryEntries));
        if (MaxPendingHistoryTasks <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPendingHistoryTasks));
        if (MaxHistoryStateBytes <= 0 || MaxHistoryTotalStateBytes < MaxHistoryStateBytes) throw new ArgumentOutOfRangeException(nameof(MaxHistoryStateBytes));
        if (!double.IsFinite(ViewportWidth) || ViewportWidth <= 0D)
            throw new ArgumentOutOfRangeException(nameof(ViewportWidth), "Viewport width must be a finite positive value.");
        if (!double.IsFinite(ViewportHeight) || ViewportHeight <= 0D)
            throw new ArgumentOutOfRangeException(nameof(ViewportHeight), "Viewport height must be a finite positive value.");
        var scripts = Scripts.ToArray();
        long length = (long)Html.Length + ReadyExpression.Length;
        foreach (string script in scripts) {
            if (script == null) throw new ArgumentException("A script cannot be null.", nameof(Scripts));
            length += script.Length;
        }
        if (length > MaxInputCharacters) throw new ArgumentException("The combined script input exceeds MaxInputCharacters.");
        HtmlRuntimeResourcePolicy.ValidateUrl(DocumentUrl);
        var policy = (ResourcePolicy ?? throw new ArgumentNullException(nameof(ResourcePolicy))).Snapshot();
        ArgumentNullException.ThrowIfNull(Resources);
        if (Resources.Count > policy.MaxRequests) throw new ArgumentException("Too many supplied resources.");
        var resources = Resources.ToArray();
        var keys = new HashSet<string>(StringComparer.Ordinal);
        long resourceBytes = 0;
        foreach (var resource in resources) {
            if (resource == null || !keys.Add(HtmlRuntimeResourcePolicy.Key(resource.Url))) throw new ArgumentException("Resources must have unique non-null URL identities.");
            if (resource.Length > policy.MaxResourceBytes || (resourceBytes += resource.Length) > policy.MaxTotalBytes)
                throw new ArgumentException("Supplied resource bytes exceed their budget.");
        }
        return new HtmlScriptRequest { Profile = Profile, Html = Html, Scripts = scripts, ReadyExpression = ReadyExpression, Timeout = Timeout,
            DocumentUrl = DocumentUrl, Resources = resources, ResourcePolicy = policy,
            SessionTimeout = SessionTimeout, PollInterval = PollInterval, MaxInputCharacters = MaxInputCharacters, MaxOutputCharacters = MaxOutputCharacters,
            MaxNodes = MaxNodes, MaxDepth = MaxDepth, MaxPendingPromiseRejections = MaxPendingPromiseRejections,
            MaxStorageCharacters = MaxStorageCharacters, MaxModuleCount = MaxModuleCount, MaxNavigations = MaxNavigations,
            MaxHistoryEntries = MaxHistoryEntries, MaxHistoryStateBytes = MaxHistoryStateBytes, MaxHistoryTotalStateBytes = MaxHistoryTotalStateBytes, MaxPendingHistoryTasks = MaxPendingHistoryTasks,
            ViewportWidth = ViewportWidth, ViewportHeight = ViewportHeight };
    }
}

/// <summary>Optional execution provider. Implementations return independent snapshots without exposing interpreter objects.</summary>
public interface IHtmlScriptRuntimeProvider {
    /// <summary>Opens a persistent trusted local document after loading HTML and executing supplied scripts.</summary>
    Task<IHtmlRuntimeSession> OpenTrustedAsync(HtmlScriptRequest request, CancellationToken cancellationToken = default);
    /// <summary>Executes trusted local content and captures it when the explicit readiness condition holds.</summary>
    Task<HtmlScriptCapture> CaptureTrustedAsync(HtmlScriptRequest request, CancellationToken cancellationToken = default);
}

/// <summary>A completed scripted-document capture. Subsequent inspection and conversion are inert.</summary>
public sealed partial class HtmlScriptCapture {
    /// <summary>Creates a capture from an independent frozen document supplied by a runtime provider.</summary>
    public HtmlScriptCapture(HtmlDocument document, string providerId, Uri? documentUrl = null, IReadOnlyList<HtmlRuntimeResource>? resources = null, Uri? baseUri = null) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentException.ThrowIfNullOrWhiteSpace(providerId);
        if (!document.IsReadOnly) throw new ArgumentException("A captured document must be frozen.", nameof(document));
        Document = document;
        ProviderId = providerId;
        DocumentUrl = HtmlRuntimeResourcePolicy.ValidateUrl(documentUrl ?? new Uri("https://officeimo.invalid/"));
        BaseUri = ResolveBaseUri(document,DocumentUrl,baseUri);
        Resources = Array.AsReadOnly((resources ?? Array.Empty<HtmlRuntimeResource>()).ToArray());
    }
    /// <summary>Frozen owned document, including structural DOM mutations and template contents.</summary>
    public HtmlDocument Document { get; }
    /// <summary>Actual runtime implementation and version used by the worker.</summary>
    public string ProviderId { get; }
    /// <summary>Current document URL, including same-document route changes.</summary>
    public Uri DocumentUrl { get; }
    /// <summary>Effective base URI at capture, including a base element frozen before a route change.</summary>
    public Uri BaseUri { get; }
    /// <summary>Immutable resource responses loaded before capture, for offline inspection or render resolution.</summary>
    public IReadOnlyList<HtmlRuntimeResource> Resources { get; }
}

/// <summary>Script execution, worker protocol or capture failed. No partial document is returned.</summary>
public sealed class HtmlScriptRuntimeException : InvalidOperationException {
    /// <summary>Creates a runtime failure with a provider-independent message.</summary>
    public HtmlScriptRuntimeException(string message) : base(message) { }
}
