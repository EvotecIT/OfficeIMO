namespace OfficeIMO.Html.Runtime.Worker;

// Fixed-field diagnostics cross the worker boundary without headers, bodies,
// credentials or script source. The client applies the caller's final redactor.
internal sealed class RuntimeDiagnostics {
    private readonly HtmlRuntimeWireTraceOptions _options;
    private readonly List<HtmlRuntimeWireEvent> _pending = new();
    private readonly object _sync = new();
    private int _recorded;
    private readonly Dictionary<string, Uri> _missingResourceUrls = new(StringComparer.Ordinal);

    internal RuntimeDiagnostics(HtmlRuntimeWireTraceOptions? options) => _options = options ?? new HtmlRuntimeWireTraceOptions();
    internal bool IncludeConsoleMessages => _options.IncludeConsoleMessages;
    internal bool IncludeFailureMessages => _options.IncludeFailureMessages;
    internal Uri[] MissingResourceUrls { get { lock (_sync) return _missingResourceUrls.Values.ToArray(); } }

    internal void RecordMissingResource(Uri url) {
        lock (_sync) _missingResourceUrls[HtmlRuntimeResourcePolicy.Key(url)] = url;
    }

    internal void ClearMissingResources() { lock (_sync) _missingResourceUrls.Clear(); }

    internal void Record(HtmlRuntimeEventKind kind, string operation, string status, DateTimeOffset started,
        TimeSpan elapsed = default, long? revision = null, string? detail = null, Uri? url = null,
        string? method = null, int? statusCode = null, long? byteCount = null, int? redirectCount = null,
        string? decision = null, string? artifactId = null) {
        if (!_options.Enabled) return;
        lock (_sync) {
            if (_recorded >= _options.MaxEvents) return;
            _recorded++;
            if (detail != null && detail.Length > _options.MaxDetailCharacters) detail = detail[.._options.MaxDetailCharacters];
            _pending.Add(new HtmlRuntimeWireEvent {
                Kind = kind, Operation = operation, Status = status, StartedUtc = started,
                ElapsedMilliseconds = elapsed.TotalMilliseconds, PageRevision = revision,
                Detail = detail, Url = _options.IncludeUrls ? url : null, Method = method,
                StatusCode = statusCode, ByteCount = byteCount, RedirectCount = redirectCount,
                Decision = decision, ArtifactId = artifactId
            });
        }
    }

    internal List<HtmlRuntimeWireEvent> Drain() {
        lock (_sync) {
            var result = new List<HtmlRuntimeWireEvent>(_pending);
            _pending.Clear();
            return result;
        }
    }
}
