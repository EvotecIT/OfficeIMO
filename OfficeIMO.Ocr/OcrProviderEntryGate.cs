using System;
using System.Diagnostics;
using System.Threading;

namespace OfficeIMO.Ocr;

/// <summary>Linearizes provider admission against caller cancellation and the total deadline.</summary>
internal sealed class OcrProviderEntryGate {
    private const int Scheduled = 0;
    private const int Started = 1;
    private const int Suppressed = 2;

    private readonly object _sync = new object();
    private readonly Stopwatch _elapsed;
    private readonly TimeSpan _timeout;
    private readonly CancellationToken _callerCancellation;
    private int _state;

    internal OcrProviderEntryGate(
        Stopwatch elapsed,
        TimeSpan timeout,
        CancellationToken callerCancellation) {
        _elapsed = elapsed ?? throw new ArgumentNullException(nameof(elapsed));
        _timeout = timeout;
        _callerCancellation = callerCancellation;
    }

    internal bool HasStarted {
        get {
            lock (_sync) return _state == Started;
        }
    }

    internal bool TryStart() {
        lock (_sync) {
            if (_state != Scheduled) return false;
            if (DeadlineOrCancellationReached()) {
                _state = Suppressed;
                return false;
            }

            _state = Started;
            // Recheck after the transition so a cancellation/deadline observed at the admission boundary
            // suppresses the call before provider-owned code is invoked.
            if (DeadlineOrCancellationReached()) {
                _state = Suppressed;
                return false;
            }

            return true;
        }
    }

    internal void SuppressIfNotStarted() {
        lock (_sync) {
            if (_state == Scheduled) _state = Suppressed;
        }
    }

    private bool DeadlineOrCancellationReached() =>
        _callerCancellation.IsCancellationRequested || _elapsed.Elapsed >= _timeout;
}
