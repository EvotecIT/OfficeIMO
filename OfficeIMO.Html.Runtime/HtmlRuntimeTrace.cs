namespace OfficeIMO.Html.Runtime;

/// <summary>Provider-neutral runtime operation category.</summary>
public enum HtmlRuntimeEventKind {
    /// <summary>Context lifecycle.</summary>
    Context,
    /// <summary>Page lifecycle.</summary>
    Page,
    /// <summary>Navigation or reload.</summary>
    Navigation,
    /// <summary>Script execution or evaluation.</summary>
    Script,
    /// <summary>Readiness or locator wait.</summary>
    Wait,
    /// <summary>Page observation.</summary>
    Observation,
    /// <summary>Structured page action.</summary>
    Action,
    /// <summary>Document capture.</summary>
    Capture,
    /// <summary>Resource request or response.</summary>
    Resource,
    /// <summary>Resource redirect.</summary>
    Redirect,
    /// <summary>Page console output.</summary>
    Console,
    /// <summary>Runtime or resource policy decision.</summary>
    Policy,
    /// <summary>Navigation or document lifecycle event.</summary>
    Lifecycle,
    /// <summary>Download request or decision.</summary>
    Download,
    /// <summary>Generated artifact evidence.</summary>
    Artifact,
    /// <summary>Failed runtime operation.</summary>
    Failure,
    /// <summary>Runtime shutdown.</summary>
    Shutdown
}

/// <summary>Caller-controlled recording and redaction for provider-neutral operation traces.</summary>
public sealed class HtmlRuntimeTraceOptions {
    /// <summary>Enables event collection.</summary>
    public bool Enabled { get; set; } = true;
    /// <summary>Maximum retained events.</summary>
    public int MaxEvents { get; set; } = 4096;
    /// <summary>Allows event URLs in traces after final redaction.</summary>
    public bool IncludeUrls { get; set; }
    /// <summary>Allows bounded page console messages in trace details.</summary>
    public bool IncludeConsoleMessages { get; set; }
    /// <summary>Allows bounded script and provider failure messages in trace details.</summary>
    public bool IncludeFailureMessages { get; set; }
    /// <summary>Maximum characters retained in one event detail after redaction.</summary>
    public int MaxDetailCharacters { get; set; } = 2048;
    /// <summary>Optional final redaction callback for event details and explicitly included URLs.</summary>
    public Func<string, string>? Redactor { get; set; }

    /// <summary>Validates and returns a detached trace configuration.</summary>
    public HtmlRuntimeTraceOptions Snapshot() {
        if (MaxEvents <= 0 || MaxEvents > 100_000) throw new ArgumentOutOfRangeException(nameof(MaxEvents));
        if (MaxDetailCharacters <= 0 || MaxDetailCharacters > 64 * 1024) throw new ArgumentOutOfRangeException(nameof(MaxDetailCharacters));
        return new HtmlRuntimeTraceOptions {
            Enabled = Enabled, MaxEvents = MaxEvents, IncludeUrls = IncludeUrls,
            IncludeConsoleMessages = IncludeConsoleMessages,
            IncludeFailureMessages = IncludeFailureMessages,
            MaxDetailCharacters = MaxDetailCharacters,
            Redactor = Redactor
        };
    }
}

/// <summary>One bounded provider-neutral runtime operation.</summary>
public sealed class HtmlRuntimeEvent {
    /// <summary>Monotonic event sequence within the page trace.</summary>
    public long Sequence { get; init; }
    /// <summary>Operation category.</summary>
    public HtmlRuntimeEventKind Kind { get; init; }
    /// <summary>Stable operation name.</summary>
    public string Operation { get; init; } = string.Empty;
    /// <summary>Success or failure status.</summary>
    public string Status { get; init; } = string.Empty;
    /// <summary>UTC operation start time.</summary>
    public DateTimeOffset StartedUtc { get; init; }
    /// <summary>Elapsed wall-clock milliseconds.</summary>
    public double ElapsedMilliseconds { get; init; }
    /// <summary>Owning context identity.</summary>
    public string ContextId { get; init; } = string.Empty;
    /// <summary>Target page identity.</summary>
    public string PageId { get; init; } = string.Empty;
    /// <summary>Page revision returned by the worker, when available.</summary>
    public long? PageRevision { get; init; }
    /// <summary>Bounded redacted detail.</summary>
    public string? Detail { get; init; }
    /// <summary>Request URL when URL recording was explicitly enabled.</summary>
    public Uri? Url { get; init; }
    /// <summary>HTTP method, when applicable. Header and body values are never recorded.</summary>
    public string? Method { get; init; }
    /// <summary>HTTP response status, when applicable.</summary>
    public int? StatusCode { get; init; }
    /// <summary>Response or artifact byte count, when known.</summary>
    public long? ByteCount { get; init; }
    /// <summary>Redirect count, when applicable.</summary>
    public int? RedirectCount { get; init; }
    /// <summary>Stable policy outcome such as allowed, blocked or unsupported.</summary>
    public string? Decision { get; init; }
    /// <summary>Content-addressed artifact identity, when an operation produced one.</summary>
    public string? ArtifactId { get; init; }
}

/// <summary>An immutable trace snapshot.</summary>
public sealed class HtmlRuntimeTrace {
    /// <summary>Provider which served the page.</summary>
    public string ProviderId { get; init; } = string.Empty;
    /// <summary>Owning context identity.</summary>
    public string ContextId { get; init; } = string.Empty;
    /// <summary>Target page identity.</summary>
    public string PageId { get; init; } = string.Empty;
    /// <summary>Whether the event limit discarded later entries.</summary>
    public bool IsTruncated { get; init; }
    /// <summary>Immutable events in sequence order.</summary>
    public IReadOnlyList<HtmlRuntimeEvent> Events { get; init; } = Array.Empty<HtmlRuntimeEvent>();
}

internal sealed class HtmlRuntimeTraceCollector {
    private readonly HtmlRuntimeTraceOptions _options;
    private readonly List<HtmlRuntimeEvent> _events = new();
    private readonly object _sync = new();
    private long _sequence;
    private bool _truncated;

    internal HtmlRuntimeTraceCollector(HtmlRuntimeTraceOptions options) => _options = options.Snapshot();
    internal bool IncludeUrls => _options.IncludeUrls;
    internal HtmlRuntimeWireTraceOptions WireOptions => new() {
        Enabled = _options.Enabled,
        MaxEvents = _options.MaxEvents,
        IncludeUrls = _options.IncludeUrls,
        IncludeConsoleMessages = _options.IncludeConsoleMessages,
        IncludeFailureMessages = _options.IncludeFailureMessages,
        MaxDetailCharacters = _options.MaxDetailCharacters
    };

    internal void Add(HtmlRuntimeEventKind kind, string operation, string status, DateTimeOffset started, TimeSpan elapsed,
        string contextId, string pageId, long? revision = null, string? detail = null,
        Uri? url = null, string? method = null, int? statusCode = null, long? byteCount = null,
        int? redirectCount = null, string? decision = null, string? artifactId = null) {
        if (!_options.Enabled) return;
        detail = Redact(detail);
        lock (_sync) {
            if (_events.Count >= _options.MaxEvents) { _truncated = true; return; }
            _events.Add(new HtmlRuntimeEvent {
                Sequence = ++_sequence,
                Kind = kind,
                Operation = operation,
                Status = status,
                StartedUtc = started,
                ElapsedMilliseconds = elapsed.TotalMilliseconds,
                ContextId = contextId,
                PageId = pageId,
                PageRevision = revision,
                Detail = detail,
                Url = _options.IncludeUrls ? RedactUrl(url) : null,
                Method = method,
                StatusCode = statusCode,
                ByteCount = byteCount,
                RedirectCount = redirectCount,
                Decision = decision,
                ArtifactId = artifactId
            });
        }
    }

    internal void Add(HtmlRuntimeWireEvent item, string contextId, string pageId) {
        string? detail = item.Kind switch {
            HtmlRuntimeEventKind.Console when !_options.IncludeConsoleMessages => null,
            HtmlRuntimeEventKind.Failure when !_options.IncludeFailureMessages => null,
            _ => item.Detail
        };
        Add(item.Kind, item.Operation, item.Status, item.StartedUtc, TimeSpan.FromMilliseconds(item.ElapsedMilliseconds),
            contextId, pageId, item.PageRevision, detail, item.Url, item.Method, item.StatusCode, item.ByteCount,
            item.RedirectCount, item.Decision, item.ArtifactId);
    }

    internal HtmlRuntimeTrace Snapshot(string providerId, string contextId, string pageId) {
        lock (_sync) return new HtmlRuntimeTrace {
            ProviderId = providerId,
            ContextId = contextId,
            PageId = pageId,
            IsTruncated = _truncated,
            Events = Array.AsReadOnly(_events.ToArray())
        };
    }

    private string? Redact(string? value) {
        if (value == null) return null;
        string? redacted = _options.Redactor == null ? value : _options.Redactor(value);
        if (redacted == null) return null;
        return redacted.Length <= _options.MaxDetailCharacters ? redacted : redacted[.._options.MaxDetailCharacters];
    }

    private Uri? RedactUrl(Uri? value) {
        if (value == null) return null;
        string? redacted = Redact(value.AbsoluteUri);
        return Uri.TryCreate(redacted, UriKind.Absolute, out Uri? result) ? result : null;
    }
}
