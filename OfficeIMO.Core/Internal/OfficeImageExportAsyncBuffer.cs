using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing.Internal;

/// <summary>
/// Provides a one-item, backpressured bridge from synchronous render producers to
/// asynchronous consumers without retaining a whole batch of encoded payloads.
/// </summary>
internal sealed class OfficeImageExportAsyncBuffer : IDisposable {
    private readonly object _gate = new object();
    private readonly Queue<OfficeImageExportResult> _queue = new Queue<OfficeImageExportResult>(1);
    private readonly SemaphoreSlim _items = new SemaphoreSlim(0);
    private readonly SemaphoreSlim _slots = new SemaphoreSlim(1);
    private bool _completed;
    private ExceptionDispatchInfo? _error;

    internal void Add(OfficeImageExportResult result, CancellationToken cancellationToken) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        _slots.Wait(cancellationToken);
        lock (_gate) {
            if (_completed) {
                _slots.Release();
                throw new InvalidOperationException("The image-export buffer is already complete.");
            }
            _queue.Enqueue(result);
        }
        _items.Release();
    }

    internal void Complete(Exception? error = null) {
        lock (_gate) {
            if (_completed) return;
            _completed = true;
            if (error != null) _error = ExceptionDispatchInfo.Capture(error);
        }
        _items.Release();
    }

    internal async Task<OfficeImageExportResult?> ReadAsync(CancellationToken cancellationToken) {
        while (true) {
            await _items.WaitAsync(cancellationToken).ConfigureAwait(false);
            OfficeImageExportResult? result = null;
            ExceptionDispatchInfo? error = null;
            bool completed;
            lock (_gate) {
                if (_queue.Count > 0) {
                    result = _queue.Dequeue();
                } else {
                    completed = _completed;
                    error = _error;
                    if (!completed) continue;
                }
            }
            if (result != null) {
                _slots.Release();
                return result;
            }
            error?.Throw();
            return null;
        }
    }

    public void Dispose() {
        _items.Dispose();
        _slots.Dispose();
    }
}
