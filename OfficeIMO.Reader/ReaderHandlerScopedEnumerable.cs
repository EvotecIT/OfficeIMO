using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

internal sealed class ReaderHandlerScopedEnumerable<T> : IEnumerable<T> {
    private readonly ReaderHandlerRegistrySnapshot _handlers;
    private readonly IEnumerable<T> _source;

    public ReaderHandlerScopedEnumerable(ReaderHandlerRegistrySnapshot handlers, IEnumerable<T> source) {
        _handlers = handlers ?? throw new ArgumentNullException(nameof(handlers));
        _source = source ?? throw new ArgumentNullException(nameof(source));
    }

    public IEnumerator<T> GetEnumerator() {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return new ReaderHandlerScopedEnumerator<T>(_handlers, _source.GetEnumerator());
        }
    }

    IEnumerator IEnumerable.GetEnumerator() {
        return GetEnumerator();
    }
}

internal sealed class ReaderHandlerScopedEnumerator<T> : IEnumerator<T> {
    private readonly ReaderHandlerRegistrySnapshot _handlers;
    private readonly IEnumerator<T> _inner;

    public ReaderHandlerScopedEnumerator(ReaderHandlerRegistrySnapshot handlers, IEnumerator<T> inner) {
        _handlers = handlers ?? throw new ArgumentNullException(nameof(handlers));
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
    }

    public T Current {
        get {
            using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
                return _inner.Current;
            }
        }
    }

    object IEnumerator.Current => Current!;

    public bool MoveNext() {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return _inner.MoveNext();
        }
    }

    public void Reset() {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            _inner.Reset();
        }
    }

    public void Dispose() {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            _inner.Dispose();
        }
    }
}
