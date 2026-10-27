using System.Diagnostics.CodeAnalysis;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel;

/// <summary>Writes XLSX worksheets as rows arrive, retaining schemas and workbook metadata rather than source rows.</summary>
/// <remarks>Each table returns its observed row count and occupied range. Dispose completes the package; file destinations are published only after every table succeeds.</remarks>
public sealed class ExcelWorkbookWriter : IDisposable {
    private readonly ExcelDocument.StreamingWorkbook _workbook;

    internal ExcelWorkbookWriter(ExcelDocument.StreamingWorkbook workbook) => _workbook = workbook;

    /// <summary>Writes one worksheet and its table from the readable scalar properties of <typeparamref name="T"/>.</summary>
    public ExcelDataSetImportResult WriteTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        string name, IEnumerable<T> rows, CancellationToken cancellationToken = default) {
        try {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            var projection = ExcelSheet.CreateStreamingRowProjection<T>();
            return WriteTable(name, rows, projection.Headers, projection.WriteRow, cancellationToken);
        } catch { _workbook.Abort(); throw; }
    }

    /// <summary>Writes one worksheet from an asynchronous sequence, awaiting and releasing one row at a time.</summary>
    public Task<ExcelDataSetImportResult> WriteTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        string name, IAsyncEnumerable<T> rows, CancellationToken cancellationToken = default) {
        try {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            var projection = ExcelSheet.CreateStreamingRowProjection<T>();
            return WriteTable(name, rows, projection.Headers, projection.WriteRow, cancellationToken);
        } catch { _workbook.Abort(); throw; }
    }

    /// <summary>Writes a single-pass sequence using an explicit column schema and a row callback.</summary>
    public ExcelDataSetImportResult WriteTable<T>(string name, IEnumerable<T> rows, IReadOnlyList<string> headers,
        Action<ExcelDocument.ExcelTabularRowWriter, T> writeRow, CancellationToken cancellationToken = default) {
        try {
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            return _workbook.WriteTable(name, new SynchronousRows<T>(rows), headers, writeRow, cancellationToken).GetAwaiter().GetResult();
        } catch { _workbook.Abort(); throw; }
    }

    /// <summary>Writes an asynchronous single-pass sequence using an explicit schema and a row callback.</summary>
    public Task<ExcelDataSetImportResult> WriteTable<T>(string name, IAsyncEnumerable<T> rows, IReadOnlyList<string> headers,
        Action<ExcelDocument.ExcelTabularRowWriter, T> writeRow, CancellationToken cancellationToken = default) =>
        _workbook.WriteTable(name, rows, headers, writeRow, cancellationToken);

    /// <summary>Writes the rows and schema of an existing data table without copying its rows.</summary>
    public ExcelDataSetImportResult WriteTable(string name, DataTable table, CancellationToken cancellationToken = default) {
        try {
            if (table == null) throw new ArgumentNullException(nameof(table));
            string[] headers = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
            return WriteTable(name, table.Rows.Cast<DataRow>(), headers, (writer, row) => {
                for (int column = 0; column < headers.Length; column++) writer.Write(row[column]);
            }, cancellationToken);
        } catch { _workbook.Abort(); throw; }
    }

    /// <summary>Writes a forward-only reader in one pass. The caller retains ownership of the reader.</summary>
    public ExcelDataSetImportResult WriteTable(string name, IDataReader reader, CancellationToken cancellationToken = default) {
        try {
            if (reader == null) throw new ArgumentNullException(nameof(reader));
            string[] headers = ExcelSheet.BuildReaderHeaders(reader);
            return WriteTable(name, ReadRows(reader), headers, (writer, row) => {
                for (int column = 0; column < headers.Length; column++) writer.Write(row.GetValue(column));
            }, cancellationToken);
        } catch { _workbook.Abort(); throw; }
    }

    private static IEnumerable<IDataRecord> ReadRows(IDataReader reader) {
        while (reader.Read()) yield return reader;
    }

    /// <summary>Completes and publishes a successful workbook, or discards a failed file export.</summary>
    public void Dispose() => _workbook.Dispose();

    private sealed class SynchronousRows<T>(IEnumerable<T> source) : IAsyncEnumerable<T> {
        public IAsyncEnumerator<T> GetAsyncEnumerator(CancellationToken cancellationToken = default) =>
            new Enumerator(source.GetEnumerator(), cancellationToken);

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
