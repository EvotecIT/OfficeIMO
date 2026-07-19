using System;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static readonly AsyncLocal<ReaderHandlerRegistrySnapshot?> ActiveHandlerRegistry = new AsyncLocal<ReaderHandlerRegistrySnapshot?>();

    internal static IDisposable UseHandlerRegistry(ReaderHandlerRegistrySnapshot snapshot) {
        if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));

        ReaderHandlerRegistrySnapshot? previous = ActiveHandlerRegistry.Value;
        ActiveHandlerRegistry.Value = snapshot;
        return new ReaderHandlerRegistryScope(previous);
    }

    private static ReaderHandlerRegistrySnapshot GetActiveHandlerRegistry() {
        return ActiveHandlerRegistry.Value ?? EmptyHandlerRegistry;
    }

    private sealed class ReaderHandlerRegistryScope : IDisposable {
        private readonly ReaderHandlerRegistrySnapshot? _previous;
        private bool _disposed;

        public ReaderHandlerRegistryScope(ReaderHandlerRegistrySnapshot? previous) {
            _previous = previous;
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }

            ActiveHandlerRegistry.Value = _previous;
            _disposed = true;
        }
    }
}
