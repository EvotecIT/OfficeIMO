using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Owns one linked caller/deadline cancellation scope for a shared image render.</summary>
internal sealed class OfficeImageExportExecutionScope : IDisposable {
    private readonly CancellationToken _callerCancellationToken;
    private readonly TimeSpan _timeout;
    private readonly CancellationTokenSource? _timeoutCancellation;
    private readonly CancellationTokenSource? _linkedCancellation;

    private OfficeImageExportExecutionScope(TimeSpan timeout, CancellationToken callerCancellationToken) {
        _timeout = timeout;
        _callerCancellationToken = callerCancellationToken;
        if (timeout == System.Threading.Timeout.InfiniteTimeSpan) {
            Token = callerCancellationToken;
            return;
        }

        _timeoutCancellation = new CancellationTokenSource();
        _timeoutCancellation.CancelAfter(timeout);
        _linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
            callerCancellationToken,
            _timeoutCancellation.Token);
        Token = _linkedCancellation.Token;
    }

    internal CancellationToken Token { get; }

    internal static OfficeImageExportExecutionScope Start(
        TimeSpan timeout,
        CancellationToken callerCancellationToken) =>
        new OfficeImageExportExecutionScope(timeout, callerCancellationToken);

    internal bool IsTimeoutCancellation(OperationCanceledException exception) =>
        _timeoutCancellation != null &&
        _timeoutCancellation.IsCancellationRequested &&
        !_callerCancellationToken.IsCancellationRequested &&
        (exception.CancellationToken == Token || Token.IsCancellationRequested);

    internal OfficeImageExportTimeoutException CreateTimeoutException(OperationCanceledException exception) =>
        new OfficeImageExportTimeoutException(_timeout, exception);

    public void Dispose() {
        _linkedCancellation?.Dispose();
        _timeoutCancellation?.Dispose();
    }
}
