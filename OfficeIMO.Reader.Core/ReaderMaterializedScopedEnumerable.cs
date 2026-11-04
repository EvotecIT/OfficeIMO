using System;
using System.Collections;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>Defers an eager read while restoring its frozen registry and inherited operation budget.</summary>
internal sealed class ReaderMaterializedScopedEnumerable<T> : IEnumerable<T> {
    private readonly ReaderHandlerRegistrySnapshot _handlers;
    private readonly Func<IEnumerable<T>> _read;
    private readonly ReaderOptions _options;
    private readonly CancellationToken _cancellationToken;

    internal ReaderMaterializedScopedEnumerable(ReaderHandlerRegistrySnapshot handlers,
        Func<IEnumerable<T>> read, ReaderOptions options, CancellationToken cancellationToken) {
        _handlers = handlers;
        _read = read;
        _options = DocumentReaderEngine.NormalizeOptions(options);
        _cancellationToken = cancellationToken;
    }

    public IEnumerator<T> GetEnumerator() => Enumerate(ReaderReadScope.CreateDetached(_options)).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private IEnumerable<T> Enumerate(ReaderReadScope readScope) {
        _cancellationToken.ThrowIfCancellationRequested();
        IEnumerable<T> items;
        using (ReaderReadScope.Use(readScope))
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            try { items = _read(); }
            finally { readScope.ClearNestedResults(); }
        }
        // The factory completes all source work and guards before returning its inert chunks.
        foreach (T item in items) {
            _cancellationToken.ThrowIfCancellationRequested();
            yield return item;
        }
    }
}
