using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Runtime;

/// <summary>Runs trusted scripted documents in disposable worker processes with bounded commands and capture.</summary>
/// <remarks>Process separation permits termination of runaway scripts. It is not an OS sandbox for hostile code.</remarks>
public sealed class HtmlProcessRuntimeProvider : IHtmlScriptRuntimeProvider {
    private readonly string _workerPath;
    private readonly string _dotnetExecutable;
    private readonly IHtmlDomServices _services;

    /// <summary>Uses a deployed OfficeIMO runtime worker DLL and the chosen inert DOM query/serialization services.</summary>
    public HtmlProcessRuntimeProvider(string workerAssemblyPath, IHtmlDomServices domServices, string dotnetExecutable = "dotnet") {
        _workerPath = Path.GetFullPath(workerAssemblyPath ?? throw new ArgumentNullException(nameof(workerAssemblyPath)));
        if (!File.Exists(_workerPath)) throw new FileNotFoundException("The runtime worker is not deployed.", _workerPath);
        _services = domServices ?? throw new ArgumentNullException(nameof(domServices));
        _dotnetExecutable = dotnetExecutable ?? throw new ArgumentNullException(nameof(dotnetExecutable));
    }

    /// <inheritdoc />
    public async Task<IHtmlRuntimeSession> OpenTrustedAsync(HtmlScriptRequest request, CancellationToken cancellationToken = default) {
        HtmlScriptRequest input = (request ?? throw new ArgumentNullException(nameof(request))).Snapshot();
        cancellationToken.ThrowIfCancellationRequested();
        var session = new HtmlProcessRuntimeSession(_workerPath, _dotnetExecutable, _services, input);
        try { await session.OpenAsync(cancellationToken).ConfigureAwait(false); return session; }
        catch { await session.DisposeAsync().ConfigureAwait(false); throw; }
    }

    /// <inheritdoc />
    public async Task<HtmlScriptCapture> CaptureTrustedAsync(HtmlScriptRequest request, CancellationToken cancellationToken = default) {
        HtmlScriptRequest input = (request ?? throw new ArgumentNullException(nameof(request))).Snapshot();
        using var deadline = new CancellationTokenSource(input.Timeout);
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, deadline.Token);
        try {
            await using IHtmlRuntimeSession session = await OpenTrustedAsync(input, operation.Token).ConfigureAwait(false);
            return await session.CaptureAsync(cancellationToken: operation.Token).ConfigureAwait(false);
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            cancellationToken.ThrowIfCancellationRequested();
            throw new TimeoutException("The scripted document did not complete within its execution deadline.");
        }
    }
}
