using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

internal sealed class ReaderHandlerScopedEnumerable<T> : IEnumerable<T> {
    private readonly ReaderHandlerRegistrySnapshot _handlers;
    private readonly IEnumerable<T> _source;
    private readonly ReaderOptions? _options;

    public ReaderHandlerScopedEnumerable(ReaderHandlerRegistrySnapshot handlers, IEnumerable<T> source, ReaderOptions? options = null) {
        _handlers = handlers ?? throw new ArgumentNullException(nameof(handlers));
        _source = source ?? throw new ArgumentNullException(nameof(source));
        _options = options == null ? null : DocumentReaderEngine.NormalizeOptions(options);
    }

    public IEnumerator<T> GetEnumerator() {
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            return new ReaderHandlerScopedEnumerator<T>(_handlers, _source.GetEnumerator(), ReaderReadScope.CreateDetached(_options));
        }
    }

    IEnumerator IEnumerable.GetEnumerator() {
        return GetEnumerator();
    }
}

internal sealed class ReaderHandlerScopedEnumerator<T> : IEnumerator<T> {
    private readonly ReaderHandlerRegistrySnapshot _handlers;
    private readonly IEnumerator<T> _inner;
    private readonly ReaderReadScope _readScope;

    public ReaderHandlerScopedEnumerator(ReaderHandlerRegistrySnapshot handlers, IEnumerator<T> inner, ReaderReadScope readScope) {
        _handlers = handlers ?? throw new ArgumentNullException(nameof(handlers));
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _readScope = readScope;
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
        using var readScope = ReaderReadScope.Use(_readScope);
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
        using var readScope = ReaderReadScope.Use(_readScope);
        using (DocumentReaderEngine.UseHandlerRegistry(_handlers)) {
            _inner.Dispose();
        }
    }
}
