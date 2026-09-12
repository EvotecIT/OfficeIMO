using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Runtime;

internal sealed class HtmlProcessRuntimeSession : IHtmlRuntimeSession {
    private readonly Process _process;
    private readonly HtmlScriptRequest _options;
    private readonly IHtmlDomServices _services;
    private readonly SemaphoreSlim _commands = new(1, 1);
    private readonly CancellationTokenSource _stopped = new();
    private readonly CancellationTokenSource _lifetime;
    private readonly CancellationTokenRegistration _lifetimeStop;
    private readonly Task _stderr;
    private readonly object _disposeSync = new();
    private Task? _disposeTask;
    private Exception? _terminationError;
    private long _commandId;
    private int _terminal;
    private int _disposed;

    internal HtmlProcessRuntimeSession(string workerPath, string dotnetExecutable, IHtmlDomServices services, HtmlScriptRequest options) {
        _options = options;
        _services = services;
        var start = new ProcessStartInfo(dotnetExecutable) { UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true, StandardErrorEncoding = Encoding.UTF8 };
        start.ArgumentList.Add(workerPath);
        _process = new Process { StartInfo = start };
        try { if (!_process.Start()) throw new HtmlScriptRuntimeException("The runtime worker could not start."); }
        catch { _process.Dispose(); _stopped.Dispose(); throw; }
        _lifetime = new CancellationTokenSource(options.SessionTimeout);
        _lifetimeStop = _lifetime.Token.Register(Stop);
        _stderr = DrainErrorsAsync();
    }

    internal Task OpenAsync(CancellationToken token) => SendAsync(new HtmlRuntimeCommand { Kind = "open", Request = _options }, token);

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
            return new HtmlScriptCapture(document.Materialize(_services, _options, token), document.ProviderId);
        }, cancellationToken);

    private HtmlRuntimeCommand Command(string kind, string script) {
        ArgumentNullException.ThrowIfNull(script);
        if (script.Length > _options.MaxInputCharacters) throw new ArgumentException("The script exceeds MaxInputCharacters.", nameof(script));
        return new HtmlRuntimeCommand { Kind = kind, Script = script };
    }

    private Task<HtmlRuntimeResponse> SendAsync(HtmlRuntimeCommand command, CancellationToken token) => SendAsync(command, (response, _) => response, token);

    private async Task<T> SendAsync<T>(HtmlRuntimeCommand command, Func<HtmlRuntimeResponse, CancellationToken, T> convert, CancellationToken token) {
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
            if (response.Error != null) throw new HtmlScriptRuntimeException(response.Error);
            T result = convert(response, operation.Token);
            operation.Token.ThrowIfCancellationRequested();
            return result;
        } catch (Exception) {
            Stop();
            await WaitForTerminationAsync().ConfigureAwait(false);
            ThrowCancellation(token, timeout);
            throw;
        } finally { _commands.Release(); }
    }

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
        try { await _process.WaitForExitAsync(deadline.Token).ConfigureAwait(false); }
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
            _process.Dispose();
            _commands.Release();
        }
    }
}
