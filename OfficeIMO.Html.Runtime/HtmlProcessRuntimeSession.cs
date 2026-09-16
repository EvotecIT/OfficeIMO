using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Runtime;

internal sealed class HtmlProcessRuntimeSession : IHtmlRuntimePage {
    private readonly Process _process;
    private readonly HtmlOciWorkerLease? _ociLease;
    private readonly HtmlScriptRequest _options;
    private readonly IHtmlDomServices _services;
    private readonly SemaphoreSlim _commands = new(1, 1);
    private readonly CancellationTokenSource _stopped = new();
    private readonly CancellationTokenSource _lifetime;
    private readonly CancellationTokenRegistration _lifetimeStop;
    private readonly Task _stderr;
    private readonly object _disposeSync = new();
    private readonly HtmlRuntimeTraceCollector _trace;
    private readonly Action? _disposedCallback;
    private Task? _disposeTask;
    private Exception? _terminationError;
    private long _commandId;
    private int _terminal;
    private int _disposed;

    internal HtmlProcessRuntimeSession(string workerPath, string dotnetExecutable, IHtmlDomServices services, HtmlScriptRequest options,
        string contextId, string pageId, HtmlRuntimeProviderDescriptor provider, HtmlRuntimeTraceOptions traceOptions, Action? disposedCallback)
        : this(StartTrustedWorker(workerPath, dotnetExecutable), null, services, options, contextId, pageId, provider, traceOptions, disposedCallback) { }

    internal HtmlProcessRuntimeSession(HtmlOciWorkerLease lease, IHtmlDomServices services, HtmlScriptRequest options,
        string contextId, string pageId, HtmlRuntimeProviderDescriptor provider, HtmlRuntimeTraceOptions traceOptions, Action? disposedCallback)
        : this(lease.Process, lease, services, options, contextId, pageId, provider, traceOptions, disposedCallback) { }

    private HtmlProcessRuntimeSession(Process process, HtmlOciWorkerLease? ociLease, IHtmlDomServices services, HtmlScriptRequest options,
        string contextId, string pageId, HtmlRuntimeProviderDescriptor provider, HtmlRuntimeTraceOptions traceOptions, Action? disposedCallback) {
        _options = options;
        _services = services;
        ContextId = contextId;
        Id = pageId;
        Provider = provider;
        _trace = new HtmlRuntimeTraceCollector(traceOptions);
        _disposedCallback = disposedCallback;
        _process = process;
        _ociLease = ociLease;
        _lifetime = new CancellationTokenSource(options.SessionTimeout);
        _lifetimeStop = _lifetime.Token.Register(Stop);
        _stderr = DrainErrorsAsync();
    }

    private static Process StartTrustedWorker(string workerPath, string dotnetExecutable) {
        var start = new ProcessStartInfo(dotnetExecutable) { UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true, StandardErrorEncoding = Encoding.UTF8 };
        start.ArgumentList.Add(workerPath);
        var process = new Process { StartInfo = start };
        try { if (!process.Start()) throw new HtmlScriptRuntimeException("The runtime worker could not start."); return process; }
        catch { process.Dispose(); throw; }
    }

    public string Id { get; }
    public string ContextId { get; }
    public HtmlRuntimeProviderDescriptor Provider { get; }

    internal Task OpenAsync(CancellationToken token) => SendAsync(new HtmlRuntimeCommand {
        Kind = "open", Request = _options, ContextId = ContextId, PageId = Id, Trace = _trace.WireOptions
    }, token);

    public Task NavigateAsync(Uri url, bool replaceHistoryEntry = false, CancellationToken cancellationToken = default) {
        EnsureWebApplicationProfile();
        var command = Command("navigate", HtmlRuntimeResourcePolicy.ValidateUrl(url).AbsoluteUri);
        command.ReplaceHistoryEntry = replaceHistoryEntry;
        return SendAsync(command, cancellationToken);
    }

    public Task ReloadAsync(CancellationToken cancellationToken = default) {
        EnsureWebApplicationProfile();
        return SendAsync(Command("reload", ""), cancellationToken);
    }

    public Task<HtmlAutomationResult> AutomateAsync(HtmlAutomationRequest request, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        var snapshot = request.Snapshot(_options.MaxInputCharacters);
        return SendAsync(new HtmlRuntimeCommand { Kind = "automation", Automation = snapshot },
            (response, _) => response.Automation ?? throw new HtmlScriptRuntimeException("The worker returned no automation result."), cancellationToken);
    }

    public Task<HtmlPageObservation> ObserveAsync(HtmlPageObservationRequest? request = null, CancellationToken cancellationToken = default) {
        HtmlPageObservationRequest snapshot = (request ?? new HtmlPageObservationRequest()).Snapshot(_options.MaxOutputCharacters);
        if (snapshot.IncludeScreenshotReference && !Provider.Supports(HtmlRuntimeCapabilityIds.ScreenshotObservation))
            throw new NotSupportedException("This runtime provider does not expose screenshot observation artifacts.");
        return SendAsync(new HtmlRuntimeCommand { Kind = "observe", Observation = snapshot, ContextId = ContextId, PageId = Id },
            (response, _) => response.Observation ?? throw new HtmlScriptRuntimeException("The worker returned no page observation."), cancellationToken);
    }

    public HtmlRuntimeTrace GetTrace() => _trace.Snapshot(Provider.Id, ContextId, Id);

    public Task ExecuteAsync(string script, CancellationToken cancellationToken = default) => SendAsync(Command("execute", script), cancellationToken);

    public Task<JsonElement> EvaluateAsync(string expression, CancellationToken cancellationToken = default) =>
        SendAsync(Command("evaluate", expression), (response, _) => {
            try {
                using JsonDocument value = JsonDocument.Parse(response.ValueJson ?? throw new HtmlScriptRuntimeException("The worker returned no JSON value."));
                return value.RootElement.Clone();
            } catch (JsonException error) { throw new HtmlScriptRuntimeException("The expression did not produce a supported JSON value: " + error.Message); }
        }, cancellationToken);

    public Task WaitForAsync(string expression, CancellationToken cancellationToken = default) => SendAsync(Command("wait", expression), cancellationToken);

    public Task<HtmlScriptCapture> CaptureAsync(string? readyExpression = null, CancellationToken cancellationToken = default) =>
        SendAsync(Command("capture", readyExpression ?? _options.ReadyExpression), (response, token) => {
            HtmlRuntimeWireDocument document = response.Document ?? throw new HtmlScriptRuntimeException("The worker returned no captured document.");
            if (document.Resources == null || document.Resources.Count > _options.ResourcePolicy.MaxRequests ||
                document.Resources.Any(resource => resource == null || resource.Length > _options.ResourcePolicy.MaxResourceBytes) ||
                document.Resources.Sum(resource => resource.Length) > _options.ResourcePolicy.MaxTotalBytes)
                throw new HtmlScriptRuntimeException("Captured resources exceed their budget.");
            return new HtmlScriptCapture(document.Materialize(_services, _options, token), document.ProviderId, document.DocumentUrl, document.Resources, document.BaseUri);
        }, cancellationToken);

    private HtmlRuntimeCommand Command(string kind, string script) {
        ArgumentNullException.ThrowIfNull(script);
        if (script.Length > _options.MaxInputCharacters) throw new ArgumentException("The script exceeds MaxInputCharacters.", nameof(script));
        return new HtmlRuntimeCommand { Kind = kind, Script = script };
    }

    private void EnsureWebApplicationProfile() {
        CheckAvailable();
        if (_options.Profile != HtmlRuntimeProfile.WebApplicationV1)
            throw new NotSupportedException("Cross-document navigation and reload require HtmlRuntimeProfile.WebApplicationV1.");
    }

    private Task<HtmlRuntimeResponse> SendAsync(HtmlRuntimeCommand command, CancellationToken token) => SendAsync(command, (response, _) => response, token);

    private async Task<T> SendAsync<T>(HtmlRuntimeCommand command, Func<HtmlRuntimeResponse, CancellationToken, T> convert, CancellationToken token) {
        DateTimeOffset started = DateTimeOffset.UtcNow;
        var stopwatch = Stopwatch.StartNew();
        CheckAvailable();
        using var timeout = new CancellationTokenSource(_options.Timeout);
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(token, timeout.Token, _stopped.Token, _lifetime.Token);
        // Admission is cancellable without terminating another caller's active command.
        try { await _commands.WaitAsync(operation.Token).ConfigureAwait(false); }
        catch (OperationCanceledException) { ThrowCancellation(token, timeout); CheckAvailable(); throw; }
        try {
            CheckAvailable();
            using var termination = operation.Token.Register(Stop);
            command.Id = ++_commandId;
            await HtmlRuntimeProtocol.WriteAsync(_process.StandardInput.BaseStream, command, HtmlRuntimeProtocol.MaximumRequestCharacters, operation.Token).ConfigureAwait(false);
            HtmlRuntimeResponse response = await HtmlRuntimeProtocol.ReadAsync<HtmlRuntimeResponse>(_process.StandardOutput.BaseStream, _options.MaxOutputCharacters, operation.Token).ConfigureAwait(false)
                ?? throw new HtmlScriptRuntimeException("The runtime worker exited before replying.");
            if (response.Id != command.Id) throw new HtmlScriptRuntimeException("The runtime response does not match its command.");
            foreach (HtmlRuntimeWireEvent item in response.Events ?? new()) _trace.Add(item, ContextId, Id);
            if (response.Error != null) {
                if (response.ErrorKind == "timeout") throw new TimeoutException(response.Error);
                throw new HtmlScriptRuntimeException(response.Error);
            }
            T result = convert(response, operation.Token);
            operation.Token.ThrowIfCancellationRequested();
            _trace.Add(EventKind(command.Kind), command.Kind, "success", started, stopwatch.Elapsed, ContextId, Id,
                response.PageRevision, TraceDetail(command), artifactId: result is HtmlScriptCapture capture ? capture.ArtifactManifest.Id : null);
            if (result is HtmlScriptCapture artifact) {
                _trace.Add(HtmlRuntimeEventKind.Artifact, "capture-manifest", "generated", started, stopwatch.Elapsed,
                    ContextId, Id, response.PageRevision, byteCount: artifact.ArtifactManifest.ByteCount,
                    artifactId: artifact.ArtifactManifest.Id);
            }
            return result;
        } catch (Exception error) {
            _trace.Add(HtmlRuntimeEventKind.Failure, command.Kind, error.GetType().Name, started, stopwatch.Elapsed,
                ContextId, Id, null, _trace.WireOptions.IncludeFailureMessages ? error.Message : null);
            Stop();
            await WaitForTerminationAsync().ConfigureAwait(false);
            ThrowCancellation(token, timeout);
            throw;
        } finally { _commands.Release(); }
    }

    private HtmlRuntimeEventKind EventKind(string kind) => kind switch {
        "open" => HtmlRuntimeEventKind.Page,
        "navigate" or "reload" => HtmlRuntimeEventKind.Navigation,
        "execute" or "evaluate" => HtmlRuntimeEventKind.Script,
        "wait" => HtmlRuntimeEventKind.Wait,
        "observe" => HtmlRuntimeEventKind.Observation,
        "automation" => HtmlRuntimeEventKind.Action,
        "capture" => HtmlRuntimeEventKind.Capture,
        _ => HtmlRuntimeEventKind.Page
    };

    private string? TraceDetail(HtmlRuntimeCommand command) => command.Kind == "navigate" && _trace.IncludeUrls ? command.Script : null;

    private void CheckAvailable() {
        if (Volatile.Read(ref _disposed) != 0) throw new ObjectDisposedException(nameof(IHtmlRuntimeSession));
        if (_lifetime.IsCancellationRequested) throw new TimeoutException("The runtime session exceeded its lifetime.");
        if (Volatile.Read(ref _terminal) != 0) throw new HtmlScriptRuntimeException("The runtime session is terminated.");
    }

    private void ThrowCancellation(CancellationToken token, CancellationTokenSource timeout) {
        token.ThrowIfCancellationRequested();
        if (timeout.IsCancellationRequested) throw new TimeoutException("The runtime command exceeded its deadline.");
        if (_lifetime.IsCancellationRequested) throw new TimeoutException("The runtime session exceeded its lifetime.");
        if (Volatile.Read(ref _disposed) != 0) throw new ObjectDisposedException(nameof(IHtmlRuntimeSession));
    }

    private void Stop() {
        if (Interlocked.Exchange(ref _terminal, 1) != 0) return;
        _stopped.Cancel();
        if (_ociLease != null) { _ociLease.Stop(); return; }
        try { if (!_process.HasExited) _process.Kill(entireProcessTree: true); }
        catch (InvalidOperationException) { /* Already exited. */ }
        catch (Win32Exception error) { _terminationError = error; }
    }

    private async Task DrainErrorsAsync() {
        try {
            var buffer = new char[4096];
            int total = 0, read;
            while ((read = await _process.StandardError.ReadAsync(buffer, _stopped.Token).ConfigureAwait(false)) != 0) {
                total += read;
                if (total > 64 * 1024) { Stop(); return; }
            }
        } catch (OperationCanceledException) when (_stopped.IsCancellationRequested) { }
        catch { Stop(); }
    }

    private async Task WaitForTerminationAsync() {
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        try {
            if (_ociLease != null) {
                await Task.WhenAll(_process.WaitForExitAsync(deadline.Token), _ociLease.WaitForRemovalAsync()).ConfigureAwait(false);
            } else {
                await _process.WaitForExitAsync(deadline.Token).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException) { throw new HtmlScriptRuntimeException("The runtime worker could not be terminated: " + _terminationError?.Message); }
    }

    public ValueTask DisposeAsync() {
        lock (_disposeSync) return new ValueTask(_disposeTask ??= DisposeCoreAsync());
    }

    private async Task DisposeCoreAsync() {
        Interlocked.Exchange(ref _disposed, 1);
        Stop();
        await _commands.WaitAsync().ConfigureAwait(false);
        try {
            await WaitForTerminationAsync().ConfigureAwait(false);
            await _stderr.ConfigureAwait(false);
        } finally {
            _lifetimeStop.Dispose();
            _lifetime.Dispose();
            _stopped.Dispose();
            try {
                if (_ociLease != null) await _ociLease.DisposeAsync().ConfigureAwait(false);
                else _process.Dispose();
            } finally {
                _commands.Release();
                _disposedCallback?.Invoke();
            }
        }
    }
}
