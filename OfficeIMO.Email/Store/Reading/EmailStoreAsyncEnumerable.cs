namespace OfficeIMO.Email.Store;

/// <summary>
/// Dependency-free async sequence that follows the C# <c>await foreach</c> pattern on every Store target framework.
/// </summary>
/// <typeparam name="T">Element type.</typeparam>
public sealed class EmailStoreAsyncEnumerable<T> {
    private readonly Func<CancellationToken, IEnumerator<T>> _factory;
    private readonly CancellationToken _operationCancellationToken;

    internal EmailStoreAsyncEnumerable(Func<CancellationToken, IEnumerator<T>> factory,
        CancellationToken operationCancellationToken) {
        _factory = factory ?? throw new ArgumentNullException(nameof(factory));
        _operationCancellationToken = operationCancellationToken;
    }

    /// <summary>Creates a single-use async enumerator. Do not use the owning Store session concurrently.</summary>
    public EmailStoreAsyncEnumerator<T> GetAsyncEnumerator(CancellationToken cancellationToken = default) {
        CancellationToken effective;
        CancellationTokenSource? linked = null;
        if (_operationCancellationToken.CanBeCanceled && cancellationToken.CanBeCanceled &&
            _operationCancellationToken != cancellationToken) {
            linked = CancellationTokenSource.CreateLinkedTokenSource(_operationCancellationToken, cancellationToken);
            effective = linked.Token;
        } else {
            effective = _operationCancellationToken.CanBeCanceled ? _operationCancellationToken : cancellationToken;
        }
        try {
            return new EmailStoreAsyncEnumerator<T>(_factory(effective), effective, linked);
        } catch {
            linked?.Dispose();
            throw;
        }
    }
}

/// <summary>Async iterator for a dependency-free Store sequence.</summary>
/// <typeparam name="T">Element type.</typeparam>
public sealed class EmailStoreAsyncEnumerator<T> {
    private readonly IEnumerator<T> _inner;
    private readonly CancellationToken _cancellationToken;
    private readonly CancellationTokenSource? _linkedCancellation;
    private int _moveInProgress;
    private bool _disposed;

    internal EmailStoreAsyncEnumerator(IEnumerator<T> inner, CancellationToken cancellationToken,
        CancellationTokenSource? linkedCancellation) {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _cancellationToken = cancellationToken;
        _linkedCancellation = linkedCancellation;
    }

    /// <summary>Current element.</summary>
    public T Current {
        get {
            if (_disposed) throw new ObjectDisposedException(GetType().FullName);
            return _inner.Current;
        }
    }

    /// <summary>Advances the already-indexed lightweight iterator and observes cancellation.</summary>
    public Task<bool> MoveNextAsync() {
        if (_disposed) throw new ObjectDisposedException(GetType().FullName);
        if (Interlocked.Exchange(ref _moveInProgress, 1) != 0) {
            throw new InvalidOperationException("Concurrent MoveNextAsync calls are not supported.");
        }
        try {
            _cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(_inner.MoveNext());
        } finally {
            Volatile.Write(ref _moveInProgress, 0);
        }
    }

    /// <summary>Disposes the underlying synchronous iterator and any linked cancellation source.</summary>
    public Task DisposeAsync() {
        if (_disposed) return Task.CompletedTask;
        if (Volatile.Read(ref _moveInProgress) != 0) {
            throw new InvalidOperationException("DisposeAsync cannot run concurrently with MoveNextAsync.");
        }
        _disposed = true;
        _inner.Dispose();
        _linkedCancellation?.Dispose();
        return Task.CompletedTask;
    }
}
