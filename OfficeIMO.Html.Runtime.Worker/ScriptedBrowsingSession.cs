namespace OfficeIMO.Html.Runtime.Worker;

// One public session, one active document. The gate serializes realm replacement
// with individual commands; waits release it between probes so navigation can run.
internal sealed class ScriptedBrowsingSession : IAsyncDisposable {
    private readonly HtmlScriptRequest _options;
    private readonly RuntimeResourceBudget _budget;
    private readonly RuntimeBrowsingStorage _storage;
    private readonly RuntimeBrowsingHistory _history = new();
    private readonly Dictionary<int, HtmlRuntimeResource> _documentSources = new();
    private readonly SemaphoreSlim _access = new(1, 1);
    private readonly object _sync = new();
    private readonly CancellationTokenSource _lifetime = new();
    private ScriptedDocumentSession? _document;
    private RuntimeNavigation? _pending;
    private Task? _pump;
    private CancellationTokenSource? _loading;
    private Exception? _failure;
    private int _generation;
    private int _navigations;
    private int _disposed;

    private ScriptedBrowsingSession(HtmlScriptRequest options) {
        _options = options;
        _budget = new(options);
        _storage = new(options.MaxStorageCharacters);
    }

    internal static async Task<ScriptedBrowsingSession> OpenAsync(HtmlScriptRequest options, CancellationToken token) {
        var session = new ScriptedBrowsingSession(options);
        await session._access.WaitAsync(token);
        try {
            session._document = await session.OpenDocumentAsync(options, token);
        } catch { session._access.Release(); await session.DisposeAsync(); throw; }
        session._access.Release();
        try { await session.SettleAsync(token); return session; }
        catch { await session.DisposeAsync(); throw; }
    }

    private Task<ScriptedDocumentSession> OpenDocumentAsync(HtmlScriptRequest request, CancellationToken token, HtmlRuntimeResource? source = null) {
        int generation = ++_generation;
        _history.Generation = generation;
        _documentSources[generation] = source ?? HtmlRuntimeResource.FromText(request.DocumentUrl, request.Html, "text/html; charset=utf-8");
        return ScriptedDocumentSession.OpenAsync(request, _budget, _storage, _history,
            request.Profile == HtmlRuntimeProfile.WebApplicationV1 ? navigation => RequestNavigation(generation, navigation) : null, token, source);
    }

    private void RequestNavigation(int generation, RuntimeNavigation navigation) {
        HtmlRuntimeResourcePolicy.ValidateUrl(navigation.Url);
        if (!_budget.Origins.Contains(HtmlRuntimeResourcePolicy.Origin(navigation.Url)))
            throw new HtmlScriptRuntimeException("The navigation origin is not allowed.");
        lock (_sync) {
            if (_lifetime.IsCancellationRequested || generation != _generation) return;
            _pending = navigation;
            _loading?.Cancel();
            _pump ??= Task.Run(PumpAsync);
        }
    }

    private async Task PumpAsync() {
        try {
            await _access.WaitAsync(_lifetime.Token);
            try {
                while (true) {
                    RuntimeNavigation navigation;
                    CancellationTokenSource deadline;
                    lock (_sync) {
                        if (_pending == null) { _pump = null; return; }
                        navigation = _pending; _pending = null;
                        deadline = CancellationTokenSource.CreateLinkedTokenSource(_lifetime.Token);
                        deadline.CancelAfter(_options.Timeout);
                        _loading = deadline;
                    }
                    using (deadline) {
                        try { await ReplaceDocumentAsync(navigation, deadline.Token); }
                        catch (OperationCanceledException) when (!_lifetime.IsCancellationRequested && HasPendingNavigation()) { }
                        finally { lock (_sync) if (ReferenceEquals(_loading, deadline)) _loading = null; }
                    }
                }
            } finally { _access.Release(); }
        } catch (OperationCanceledException) when (_lifetime.IsCancellationRequested) {
            lock (_sync) { _pending = null; _pump = null; }
        } catch (Exception error) {
            lock (_sync) { _failure = error; _pending = null; _pump = null; }
        }
    }

    private bool HasPendingNavigation() { lock (_sync) return _pending != null; }

    private async Task ReplaceDocumentAsync(RuntimeNavigation navigation, CancellationToken token) {
        if (++_navigations > _options.MaxNavigations) throw new HtmlScriptRuntimeException("The navigation count budget was exceeded.");
        HtmlRuntimeResource? response;
        bool replayRetainedSource = TryGetRetainedSource(navigation, out response);
        if (!replayRetainedSource) {
            using var loader = new RuntimeResourceLoader(_options, _budget);
            response = await loader.LoadAsync(navigation.Url, token);
            if (response.StatusCode is 204 or 205) return;
            if (response.StatusCode == 304) throw new HtmlScriptRuntimeException("A 304 navigation response requires an HTTP cache, which this profile does not provide.");
            if (response.Headers.TryGetValue("Content-Disposition", out var disposition) && disposition.TrimStart().StartsWith("attachment", StringComparison.OrdinalIgnoreCase))
                throw new HtmlScriptRuntimeException("Navigation downloads are outside this document profile.");
            if (!response.ContentType.Split(';', 2)[0].Trim().Equals("text/html", StringComparison.OrdinalIgnoreCase))
                throw new HtmlScriptRuntimeException("Navigation requires an HTML response.");
        }
        token.ThrowIfCancellationRequested();
        _document?.Dispose();
        _document = null;
        _history.Transition = navigation;
        var request = _options.Snapshot();
        request.DocumentUrl = replayRetainedSource ? navigation.Url : response!.FinalUrl;
        request.Html = string.Empty;
        request.Scripts = Array.Empty<string>();
        _document = await OpenDocumentAsync(request, token, response);
        if (navigation.EntryIndex >= 0 && !navigation.Reload) await _document.RestoreTraversalAsync(token);
    }

    private bool TryGetRetainedSource(RuntimeNavigation navigation, out HtmlRuntimeResource? source) {
        string key = HtmlRuntimeResourcePolicy.Key(navigation.Url);
        source = null;
        if (_options.ResourcePolicy.AllowNetwork
            || navigation.DocumentId < 0
            || _options.Resources.Any(resource => HtmlRuntimeResourcePolicy.Key(resource.Url) == key)) return false;
        return _documentSources.TryGetValue(navigation.DocumentId, out source);
    }

    private async Task SettleAsync(CancellationToken token) {
        while (true) {
            Task? pump;
            lock (_sync) { if (_failure != null) throw new HtmlScriptRuntimeException(_failure.Message); pump = _pump; }
            if (pump == null) return;
            await pump.WaitAsync(token);
        }
    }

    private async Task<(T Value, int Generation)> OnDocumentAsync<T>(Func<ScriptedDocumentSession, Task<T>> operation, CancellationToken token) {
        while (true) {
            await SettleAsync(token);
            await _access.WaitAsync(token);
            try {
                if (HasPendingNavigation()) continue;
                lock (_sync) if (_failure != null) throw new HtmlScriptRuntimeException(_failure.Message);
                int generation = _generation;
                T value = await operation(_document ?? throw new HtmlScriptRuntimeException("There is no active document."));
                return (value, generation);
            } finally { _access.Release(); }
        }
    }

    internal async Task ExecuteAsync(string script, CancellationToken token) {
        await OnDocumentAsync(async document => { await document.ExecuteAsync(script, token); return true; }, token);
        await SettleAsync(token);
    }

    internal async Task NavigateAsync(string url, bool replace, CancellationToken token) {
        await OnDocumentAsync(async document => { await document.NavigateAsync(url, replace, token); return true; }, token);
        await SettleAsync(token);
    }

    internal async Task ReloadAsync(CancellationToken token) {
        await OnDocumentAsync(async document => { await document.ReloadAsync(token); return true; }, token);
        await SettleAsync(token);
    }

    internal async Task<string> EvaluateAsync(string expression, CancellationToken token) {
        var result = await OnDocumentAsync(document => document.EvaluateAsync(expression, token), token);
        await SettleAsync(token);
        return result.Value;
    }

    internal async Task<HtmlAutomationResult> AutomateAsync(HtmlAutomationRequest request, CancellationToken token) {
        bool observation = request.Action is HtmlAutomationAction.Inspect or HtmlAutomationAction.Count or HtmlAutomationAction.Wait;
        while (true) {
            var result = await OnDocumentAsync(document => document.AutomateAsync(request, token), token);
            await SettleAsync(token);
            if (observation && result.Generation != Volatile.Read(ref _generation)) continue;
            if (!request.WaitForReady || result.Value.Status is not (HtmlAutomationStatus.NotFound or HtmlAutomationStatus.NotReady)) return result.Value;
            await Task.Delay(_options.PollInterval, token);
        }
    }

    internal async Task<HtmlRuntimeWireDocument?> WaitAsync(string expression, bool capture, CancellationToken token) {
        while (true) {
            var result = await OnDocumentAsync(document => document.ProbeAsync(expression, capture, token), token);
            await SettleAsync(token);
            if (result.Generation == _generation && result.Value.Ready) return result.Value.Document;
            await Task.Delay(_options.PollInterval, token);
        }
    }

    public async ValueTask DisposeAsync() {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        _lifetime.Cancel();
        Task? pump;
        lock (_sync) { _loading?.Cancel(); pump = _pump; }
        if (pump != null) await pump;
        await _access.WaitAsync();
        try { _document?.Dispose(); _document = null; _history.Entries = Jint.Native.JsValue.Null; }
        finally { _access.Release(); }
        await _budget.StopAndWaitAsync();
        _budget.Dispose();
        _lifetime.Dispose();
        _access.Dispose();
    }
}
