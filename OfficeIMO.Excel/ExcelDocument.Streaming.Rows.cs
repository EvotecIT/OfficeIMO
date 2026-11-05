using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Uses the same callback lifecycle for synchronous sources and genuinely asynchronous rows.</summary>
    private static async Task<int> WriteStreamingRowCallbacks<T>(IAsyncEnumerable<T> rows, ExcelTabularRowWriter writer,
        Action<ExcelTabularRowWriter, T> writeRow, int maximumRows, CancellationToken ct) {
        int count = 0;
        IAsyncEnumerator<T> enumerator = rows.GetAsyncEnumerator(ct);
        try {
            while (true) {
                ct.ThrowIfCancellationRequested();
                if (!await enumerator.MoveNextAsync().ConfigureAwait(false)) break;
                ct.ThrowIfCancellationRequested();
                WriteNext(enumerator.Current);
            }
        } finally { await enumerator.DisposeAsync().ConfigureAwait(false); }
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
}
