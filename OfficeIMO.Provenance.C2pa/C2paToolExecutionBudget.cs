using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Provenance.C2pa;

/// <summary>Applies one cancellation and timeout budget across provider execution and report interpretation.</summary>
internal sealed class C2paToolExecutionBudget {
    private readonly CancellationToken _cancellationToken;
    private readonly TimeSpan _timeout;
    private readonly Stopwatch _stopwatch = Stopwatch.StartNew();

    internal C2paToolExecutionBudget(TimeSpan timeout, CancellationToken cancellationToken) {
        _timeout = timeout;
        _cancellationToken = cancellationToken;
    }

    internal void ThrowIfExceeded() {
        _cancellationToken.ThrowIfCancellationRequested();
        if (_stopwatch.Elapsed >= _timeout) {
            throw CreateTimeoutException();
        }
    }

    internal TimeSpan GetRemainingTimeout() {
        ThrowIfExceeded();
        return _timeout - _stopwatch.Elapsed;
    }

    internal T RunInterpretation<T>(Func<CancellationToken, T> interpretation) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(interpretation);
#else
        if (interpretation is null) throw new ArgumentNullException(nameof(interpretation));
#endif
        ThrowIfExceeded();
        TimeSpan remaining = _timeout - _stopwatch.Elapsed;
        using var interpretationCancellation = CancellationTokenSource.CreateLinkedTokenSource(_cancellationToken);
        Task<T> work = Task.Run(() => interpretation(interpretationCancellation.Token), CancellationToken.None);
        Task deadline = Task.Delay(remaining, interpretationCancellation.Token);
        Task completed = Task.WhenAny(work, deadline).GetAwaiter().GetResult();
        if (ReferenceEquals(completed, work)) {
            interpretationCancellation.Cancel();
            return work.GetAwaiter().GetResult();
        }

        interpretationCancellation.Cancel();
        _ = work.ContinueWith(
            static task => _ = task.Exception,
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously | TaskContinuationOptions.OnlyOnFaulted,
            TaskScheduler.Default);
        _cancellationToken.ThrowIfCancellationRequested();
        throw CreateTimeoutException();
    }

    private static TimeoutException CreateTimeoutException() =>
        new("c2patool provider execution and report interpretation exceeded the configured timeout.");
}
