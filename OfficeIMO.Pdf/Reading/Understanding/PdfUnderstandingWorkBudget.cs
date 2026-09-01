using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Shared cancellation-aware work budget for semantic reconstruction.</summary>
internal sealed class PdfUnderstandingWorkBudget {
    private readonly CancellationToken _cancellationToken;
    private readonly long _maximum;
    private long _consumed;
    private int _operationCompleted;

    internal PdfUnderstandingWorkBudget(long maximum, CancellationToken cancellationToken) {
#pragma warning disable CA1512 // ThrowIfNegativeOrZero is unavailable on every target framework.
        if (maximum <= 0) throw new ArgumentOutOfRangeException(nameof(maximum));
#pragma warning restore CA1512
        _maximum = maximum;
        _cancellationToken = cancellationToken;
        _cancellationToken.ThrowIfCancellationRequested();
    }

    internal long Maximum => _maximum;
    internal long Consumed => _consumed;
    internal CancellationToken CancellationToken => Volatile.Read(ref _operationCompleted) == 0
        ? _cancellationToken
        : CancellationToken.None;

    internal void Consume(long units = 1) {
#pragma warning disable CA1512 // ThrowIfNegativeOrZero is unavailable on every target framework.
        if (units <= 0) throw new ArgumentOutOfRangeException(nameof(units));
#pragma warning restore CA1512
        if (Volatile.Read(ref _operationCompleted) != 0) return;
        ThrowIfCancellationRequested();
        long next;
        try {
            next = checked(_consumed + units);
        } catch (OverflowException) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingWork, _maximum, long.MaxValue);
        }
        if (next > _maximum) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingWork, _maximum, next);
        }
        _consumed = next;
    }

    internal void ThrowIfCancellationRequested() {
        if (Volatile.Read(ref _operationCompleted) == 0) {
            _cancellationToken.ThrowIfCancellationRequested();
        }
    }

    /// <summary>Detaches retained lazy projections from the completed read operation's token and one-time work budget.</summary>
    internal void CompleteOperation() => Volatile.Write(ref _operationCompleted, 1);
}
