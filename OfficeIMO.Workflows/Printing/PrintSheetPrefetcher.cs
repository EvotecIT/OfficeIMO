namespace OfficeIMO.Workflows;

/// <summary>
/// Prepares a transport's sheets a few ahead on worker threads and hands them back in order, so rendering runs on several cores
/// while the transport sends the current sheet. Items are taken in increasing order; at most <c>depth</c> are prepared ahead.
/// </summary>
internal sealed class PrintSheetPrefetcher<T> : IDisposable where T : class {
    private readonly Func<int, CancellationToken, T?> _prepare;
    private readonly Task<T?>?[] _pending;
    private readonly int _depth;
    private readonly CancellationTokenSource _stop;
    private int _scheduled;

    /// <param name="count">Number of items, taken by index from 0.</param>
    /// <param name="depth">Items prepared ahead of the one being taken; 1 prepares only the next item.</param>
    /// <param name="prepare">Prepares one item, or returns null when the transport handles that item itself, such as a repeated sheet.</param>
    /// <param name="token">Stops preparation.</param>
    internal PrintSheetPrefetcher(int count, int depth, Func<int, CancellationToken, T?> prepare, CancellationToken token) {
        _prepare = prepare;
        _pending = new Task<T?>?[count];
        _depth = Math.Max(1, depth);
        _stop = CancellationTokenSource.CreateLinkedTokenSource(token);
    }

    /// <summary>
    /// How many sheets to prepare ahead: one per spare core up to four, and no more than the content byte limit holds at the
    /// expected sheet size, so a job never keeps more pixels in memory than it is allowed to compose.
    /// </summary>
    internal static int DepthFor(long sheetBytes, long maximumContentBytes) {
        long byMemory = sheetBytes > 0 ? maximumContentBytes / sheetBytes : 1;
        return (int)Math.Max(1, Math.Min(Math.Min(Environment.ProcessorCount - 1, 4), byMemory));
    }

    /// <summary>Waits for item <paramref name="index"/>, preparing it now if it was not started, and starts the items after it.</summary>
    internal T? Take(int index) {
        Schedule(index);
        Task<T?> task = _pending[index]!;
        _pending[index] = null;
        T? item = task.GetAwaiter().GetResult();
        Schedule(index + 1);
        return item;
    }

    private void Schedule(int from) {
        while (_scheduled < _pending.Length && _scheduled < from + _depth) {
            int index = _scheduled++;
            CancellationToken token = _stop.Token;
            _pending[index] = Task.Run(() => _prepare(index, token), token);
        }
    }

    /// <summary>Stops preparation and waits for running items, so no render outlives the delivery that started it.</summary>
    public void Dispose() {
        _stop.Cancel();
        foreach (Task<T?>? task in _pending) {
            if (task is null) continue;
            try { task.GetAwaiter().GetResult(); } catch { /* The delivery outcome is already being reported. */ }
        }
        _stop.Dispose();
    }
}
