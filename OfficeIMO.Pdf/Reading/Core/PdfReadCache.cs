using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Publishes one completed read result while keeping callers' waits cancellable.</summary>
internal sealed class PdfReadCache<T> {
    private readonly object _sync = new();
    private T _value = default!;
    private int _initialized;

    internal PdfReadCache() { }

    internal PdfReadCache(T value) {
        _value = value;
        _initialized = 1;
    }

    /// <summary>Retries failed or cancelled initialization; only successful results enter the cache.</summary>
    internal T GetOrCreate<TState>(TState state, Func<TState, CancellationToken, T> factory, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (Volatile.Read(ref _initialized) != 0) return _value;

        while (!Monitor.TryEnter(_sync, millisecondsTimeout: 25)) {
            cancellationToken.ThrowIfCancellationRequested();
        }
        try {
            cancellationToken.ThrowIfCancellationRequested();
            if (_initialized == 0) {
                T value = factory(state, cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                _value = value;
                Volatile.Write(ref _initialized, 1);
            }
            return _value;
        } finally {
            Monitor.Exit(_sync);
        }
    }
}
