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
    /// <summary>Allows navigation URLs in trace details.</summary>
    public bool IncludeUrls { get; set; }
    /// <summary>Optional final redaction callback for event details.</summary>
    public Func<string, string>? Redactor { get; set; }

    /// <summary>Validates and returns a detached trace configuration.</summary>
    public HtmlRuntimeTraceOptions Snapshot() {
        if (MaxEvents <= 0 || MaxEvents > 100_000) throw new ArgumentOutOfRangeException(nameof(MaxEvents));
        return new HtmlRuntimeTraceOptions { Enabled = Enabled, MaxEvents = MaxEvents, IncludeUrls = IncludeUrls, Redactor = Redactor };
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

    internal void Add(HtmlRuntimeEventKind kind, string operation, string status, DateTimeOffset started, TimeSpan elapsed,
        string contextId, string pageId, long? revision = null, string? detail = null) {
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
                Detail = detail
            });
        }
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
        string redacted = _options.Redactor?.Invoke(value) ?? value;
        return redacted.Length <= 2048 ? redacted : redacted[..2048];
    }
}
