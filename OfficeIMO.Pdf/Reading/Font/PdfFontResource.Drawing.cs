using System.Threading;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfFontResource {
    private DeferredDrawing? _deferredDrawing;
    private Lazy<string?[]?>? _builtinEncoding;

    internal bool HasDeferredDrawing => _deferredDrawing != null && _deferredDrawing.Resolved == null;
    internal bool HasBuiltinEncoding => _builtinEncoding != null;
    internal string?[]? BuiltinEncoding => _builtinEncoding?.Value;
    internal PdfFontResource? RetainedDrawingResource => _deferredDrawing?.Resolved;

    internal PdfFontResource WithDeferredDrawing(Func<CancellationToken, PdfFontResource> factory) {
        var copy = new PdfFontResource(ResourceName, this);
        copy._deferredDrawing = new DeferredDrawing(factory);
        return copy;
    }

    internal void SetBuiltinEncoding(Func<string?[]?> factory) => _builtinEncoding = new Lazy<string?[]?>(factory);

    internal void ObserveDrawingResolution(Action<PdfFontResource> resolved) {
        if (_deferredDrawing != null) _deferredDrawing.OnResolved = resolved;
    }

    /// <summary>Materializes outlines only for drawing. Aliases share successful work; failed operations can retry.</summary>
    internal PdfFontResource ForDrawing(CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfFontResource resolved = _deferredDrawing?.Get(cancellationToken) ?? this;
        return resolved.WithResourceName(ResourceName);
    }

    private sealed class DeferredDrawing {
        private readonly object _sync = new();
        private Func<CancellationToken, PdfFontResource>? _factory;
        private PdfFontResource? _resolved;
        internal Action<PdfFontResource>? OnResolved { get; set; }
        internal PdfFontResource? Resolved => Volatile.Read(ref _resolved);
        internal DeferredDrawing(Func<CancellationToken, PdfFontResource> factory) { _factory = factory; }

        internal PdfFontResource Get(CancellationToken cancellationToken) {
            lock (_sync) {
                cancellationToken.ThrowIfCancellationRequested();
                if (_resolved != null) return _resolved;
                PdfFontResource resolved = _factory!(cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                // Cache admission uses a nonforcing snapshot and may evict the owning metadata
                // entry when the materialized program exceeds the document's retention budget.
                OnResolved?.Invoke(resolved);
                _resolved = resolved;
                _factory = null;
                OnResolved = null;
                return resolved;
            }
        }
    }
}
