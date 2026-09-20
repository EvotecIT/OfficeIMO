namespace OfficeIMO.Html.Runtime.Worker;

// Fixed-field diagnostics cross the worker boundary without headers, bodies,
// credentials or script source. The client applies the caller's final redactor.
internal sealed class RuntimeDiagnostics {
    internal const string MissingFetchBudgetMessage = "Offline fetch discovery exceeds its response character budget.";
    private readonly HtmlRuntimeWireTraceOptions _options;
    private readonly int _maxMissingFetchCharacters;
    private readonly List<HtmlRuntimeWireEvent> _pending = new();
    private readonly object _sync = new();
    private int _recorded;
    private readonly Dictionary<string, Uri> _missingResourceUrls = new(StringComparer.Ordinal);
    private readonly Dictionary<string, HtmlRuntimeFetchDiscovery> _missingFetchRequests = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _missingFetchCharacters = new(StringComparer.Ordinal);
    private readonly List<string> _consumedFetchReplayIdentities = new();
    private readonly Dictionary<string, HtmlRuntimeNavigationDiscovery> _missingNavigationRequests = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _missingNavigationCharacters = new(StringComparer.Ordinal);
    private readonly List<string> _consumedNavigationReplayIdentities = new();
    private long _retainedMissingFetchCharacters = 2; // JSON array brackets.
    private bool _missingFetchBudgetExceeded;

    internal RuntimeDiagnostics(HtmlRuntimeWireTraceOptions? options, int maxResponseCharacters) {
        _options = options ?? new HtmlRuntimeWireTraceOptions();
        _maxMissingFetchCharacters = maxResponseCharacters / 2;
    }
    internal bool IncludeConsoleMessages => _options.IncludeConsoleMessages;
    internal bool IncludeFailureMessages => _options.IncludeFailureMessages;
    internal Uri[] MissingResourceUrls { get { lock (_sync) return _missingResourceUrls.Values.ToArray(); } }
    internal HtmlRuntimeFetchDiscovery[] MissingFetchRequests { get { lock (_sync) return _missingFetchRequests.Values.ToArray(); } }
    internal string[] ConsumedFetchReplayIdentities { get { lock (_sync) return _consumedFetchReplayIdentities.ToArray(); } }
    internal HtmlRuntimeNavigationDiscovery[] MissingNavigationRequests { get { lock (_sync) return _missingNavigationRequests.Values.ToArray(); } }
    internal string[] ConsumedNavigationReplayIdentities { get { lock (_sync) return _consumedNavigationReplayIdentities.ToArray(); } }
    internal bool MissingFetchBudgetExceeded { get { lock (_sync) return _missingFetchBudgetExceeded; } }

    internal void RecordMissingResource(Uri url) {
        lock (_sync) _missingResourceUrls[HtmlRuntimeResourcePolicy.Key(url)] = url;
    }

    internal void RecordMissingFetch(HtmlRuntimeFetchDiscovery request) {
        int characters = HtmlRuntimeProtocol.MeasureCharacters(request);
        lock (_sync) {
            int previous = _missingFetchCharacters.GetValueOrDefault(request.Identity);
            long separators = previous == 0 && (_missingFetchRequests.Count + _missingNavigationRequests.Count) != 0 ? 1 : 0;
            long projected = _retainedMissingFetchCharacters - previous + characters + separators;
            if (projected > _maxMissingFetchCharacters) {
                _missingFetchBudgetExceeded = true;
                throw new HtmlScriptRuntimeException(MissingFetchBudgetMessage);
            }
            _missingFetchRequests[request.Identity] = request;
            _missingFetchCharacters[request.Identity] = characters;
            _retainedMissingFetchCharacters = projected;
        }
    }

    internal void RecordConsumedFetchReplay(string identity) {
        lock (_sync) _consumedFetchReplayIdentities.Add(identity);
    }

    internal void RecordMissingNavigation(HtmlRuntimeNavigationDiscovery request) {
        int characters = HtmlRuntimeProtocol.MeasureCharacters(request);
        lock (_sync) {
            int previous = _missingNavigationCharacters.GetValueOrDefault(request.Identity);
            long separators = previous == 0 && (_missingFetchRequests.Count + _missingNavigationRequests.Count) != 0 ? 1 : 0;
            long projected = _retainedMissingFetchCharacters - previous + characters + separators;
            if (projected > _maxMissingFetchCharacters) {
                _missingFetchBudgetExceeded = true;
                throw new HtmlScriptRuntimeException(MissingFetchBudgetMessage);
            }
            _missingNavigationRequests[request.Identity] = request;
            _missingNavigationCharacters[request.Identity] = characters;
            _retainedMissingFetchCharacters = projected;
        }
    }

    internal void RecordConsumedNavigationReplay(string identity) {
        lock (_sync) _consumedNavigationReplayIdentities.Add(identity);
    }

    internal void ClearMissingResources() {
        lock (_sync) {
            _missingResourceUrls.Clear();
            _missingFetchRequests.Clear();
            _missingFetchCharacters.Clear();
            _missingNavigationRequests.Clear();
            _missingNavigationCharacters.Clear();
            _retainedMissingFetchCharacters = 2;
            _missingFetchBudgetExceeded = false;
        }
    }

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
