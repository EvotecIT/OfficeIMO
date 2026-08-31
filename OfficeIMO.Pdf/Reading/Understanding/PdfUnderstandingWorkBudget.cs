using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Shared cancellation-aware work budget for semantic reconstruction.</summary>
internal sealed class PdfUnderstandingWorkBudget {
    private readonly CancellationToken _cancellationToken;
    private readonly long _maximum;
    private long _consumed;

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
    internal CancellationToken CancellationToken => _cancellationToken;

    internal void Consume(long units = 1) {
#pragma warning disable CA1512 // ThrowIfNegativeOrZero is unavailable on every target framework.
        if (units <= 0) throw new ArgumentOutOfRangeException(nameof(units));
#pragma warning restore CA1512
        _cancellationToken.ThrowIfCancellationRequested();
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

    internal void ThrowIfCancellationRequested() => _cancellationToken.ThrowIfCancellationRequested();
}
