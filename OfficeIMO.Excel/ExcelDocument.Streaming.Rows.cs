namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Uses the same callback lifecycle for synchronous sources and genuinely asynchronous rows.</summary>
    private static async Task<int> WriteStreamingRowCallbacks<T>(IAsyncEnumerable<T> rows, ExcelTabularRowWriter writer,
        Action<ExcelTabularRowWriter, T> writeRow, int maximumRows, CancellationToken ct) {
        int count = 0;
        if (rows is SynchronousRows<T> synchronousRows) {
            using IEnumerator<T> enumerator = synchronousRows.Source.GetEnumerator();
            while (true) {
                ct.ThrowIfCancellationRequested();
                if (!enumerator.MoveNext()) break;
                ct.ThrowIfCancellationRequested();
                WriteNext(enumerator.Current);
            }
        } else {
            IAsyncEnumerator<T> enumerator = rows.GetAsyncEnumerator(ct);
            try {
                while (true) {
                    ct.ThrowIfCancellationRequested();
                    if (!await enumerator.MoveNextAsync().ConfigureAwait(false)) break;
                    ct.ThrowIfCancellationRequested();
                    WriteNext(enumerator.Current);
                }
            } finally { await enumerator.DisposeAsync().ConfigureAwait(false); }
        }
        ct.ThrowIfCancellationRequested();
        return count;

        void WriteNext(T item) {
            if (count >= maximumRows) throw new InvalidOperationException("Row export exceeds the selected format's maximum worksheet row count.");
            writer.BeginRow();
            try {
                writeRow(writer, item);
                writer.EndRow();
                count++;
            } finally { writer.ReleaseRow(); }
        }
    }

    /// <summary>Retains a synchronous source until the common table setup has completed.</summary>
    internal sealed class SynchronousRows<T>(IEnumerable<T> source) : IAsyncEnumerable<T> {
        /// <summary>Provides the original one-pass source to the synchronous row loop.</summary>
        internal IEnumerable<T> Source { get; } = source;

        public IAsyncEnumerator<T> GetAsyncEnumerator(CancellationToken cancellationToken = default) =>
            new Enumerator(Source.GetEnumerator(), cancellationToken);

        private sealed class Enumerator(IEnumerator<T> source, CancellationToken cancellationToken) : IAsyncEnumerator<T> {
            public T Current => source.Current;
            public ValueTask<bool> MoveNextAsync() {
                cancellationToken.ThrowIfCancellationRequested();
                return new ValueTask<bool>(source.MoveNext());
            }
            public ValueTask DisposeAsync() { source.Dispose(); return default; }
        }
    }
}
