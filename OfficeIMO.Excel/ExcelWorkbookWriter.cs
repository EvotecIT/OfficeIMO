using System.Diagnostics.CodeAnalysis;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel;

/// <summary>Writes Excel worksheets through one typed row API. XLSX and XLSB stream rows; XLS stages native records and shared strings.</summary>
/// <remarks>Each table returns its observed row count and occupied range. Dispose completes the package; file destinations are published only after every table succeeds. Disposing before any table was written throws <see cref="InvalidOperationException"/> and leaves the destination untouched.</remarks>
public sealed class ExcelWorkbookWriter : IDisposable {
    private readonly ExcelDocument.StreamingWorkbook _workbook;
    private ExcelIncrementalTableWriter? _activeTable;

    internal ExcelWorkbookWriter(ExcelDocument.StreamingWorkbook workbook) => _workbook = workbook;

    /// <summary>Begins a disk-spooled table whose final schema may grow as normalized rows arrive.</summary>
    public ExcelIncrementalTableWriter BeginTable(string name, CancellationToken cancellationToken = default) {
        EnsureNoActiveTable();
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        cancellationToken.ThrowIfCancellationRequested();
        return _activeTable = new ExcelIncrementalTableWriter(this, name, _workbook.MaximumTemporaryBytes, cancellationToken);
    }

    internal void AbortIncrementalTable(ExcelIncrementalTableWriter table) => _workbook.Abort();

    internal ExcelDataSetImportResult CompleteIncrementalTable(ExcelIncrementalTableWriter table, string name,
        IReadOnlyList<string> headers, IEnumerable<object?[]> rows, CancellationToken cancellationToken) {
        if (!ReferenceEquals(_activeTable, table)) throw new InvalidOperationException("This table is not active on the workbook.");
        _activeTable = null;
        return WriteTable(name, rows, headers, (writer, values) => {
            foreach (object? value in values) { if (value == null) writer.WriteBlank(); else writer.Write(value); }
        }, cancellationToken);
    }

    private void EnsureNoActiveTable() {
        if (_activeTable != null) throw new InvalidOperationException("Complete or dispose the current incremental table first.");
    }

    /// <summary>Writes one worksheet and its table from the readable scalar properties of <typeparamref name="T"/>.</summary>
    public ExcelDataSetImportResult WriteTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        string name, IEnumerable<T> rows, CancellationToken cancellationToken = default) {
        try {
            ArgumentNullException.ThrowIfNull(rows);
            var projection = ExcelSheet.CreateStreamingRowProjection<T>();
            return WriteTable(name, rows, projection.Headers, projection.WriteRow, cancellationToken);
        } catch { _workbook.Abort(); throw; }
    }

    /// <summary>Writes one worksheet from an asynchronous sequence, awaiting and releasing one row at a time.</summary>
    public Task<ExcelDataSetImportResult> WriteTable<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        string name, IAsyncEnumerable<T> rows, CancellationToken cancellationToken = default) {
        try {
            ArgumentNullException.ThrowIfNull(rows);
            var projection = ExcelSheet.CreateStreamingRowProjection<T>();
            return WriteTable(name, rows, projection.Headers, projection.WriteRow, cancellationToken);
        } catch { _workbook.Abort(); throw; }
    }

    /// <summary>Writes a single-pass sequence using an explicit column schema and a row callback.</summary>
    public ExcelDataSetImportResult WriteTable<T>(string name, IEnumerable<T> rows, IReadOnlyList<string> headers,
        Action<ExcelDocument.ExcelTabularRowWriter, T> writeRow, CancellationToken cancellationToken = default) {
        try {
            EnsureNoActiveTable();
            ArgumentNullException.ThrowIfNull(rows);
            return _workbook.WriteTable(name, new ExcelDocument.SynchronousRows<T>(rows), headers, writeRow, cancellationToken).GetAwaiter().GetResult();
        } catch { _workbook.Abort(); throw; }
    }

    /// <summary>Writes an asynchronous single-pass sequence using an explicit schema and a row callback.</summary>
    public Task<ExcelDataSetImportResult> WriteTable<T>(string name, IAsyncEnumerable<T> rows, IReadOnlyList<string> headers,
        Action<ExcelDocument.ExcelTabularRowWriter, T> writeRow, CancellationToken cancellationToken = default) =>
        WriteAsyncTable(name, rows, headers, writeRow, cancellationToken);

    private Task<ExcelDataSetImportResult> WriteAsyncTable<T>(string name, IAsyncEnumerable<T> rows, IReadOnlyList<string> headers,
        Action<ExcelDocument.ExcelTabularRowWriter, T> writeRow, CancellationToken cancellationToken) {
        EnsureNoActiveTable();
        return _workbook.WriteTable(name, rows, headers, writeRow, cancellationToken);
    }

    /// <summary>Writes the rows and schema of an existing data table without copying its rows.</summary>
    public ExcelDataSetImportResult WriteTable(string name, DataTable table, CancellationToken cancellationToken = default) {
        try {
            ArgumentNullException.ThrowIfNull(table);
            string[] headers = table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
            return WriteTable(name, table.Rows.Cast<DataRow>(), headers, (writer, row) => {
                for (int column = 0; column < headers.Length; column++) {
                    if (row.IsNull(column)) writer.WriteBlank();
                    else writer.Write(row[column]);
                }
            }, cancellationToken);
        } catch { _workbook.Abort(); throw; }
    }

    /// <summary>Writes a forward-only reader in one pass. The caller retains ownership of the reader.</summary>
    public ExcelDataSetImportResult WriteTable(string name, IDataReader reader, CancellationToken cancellationToken = default) {
        try {
            ArgumentNullException.ThrowIfNull(reader);
            string[] headers = ExcelSheet.BuildReaderHeaders(reader);
            return WriteTable(name, ReadRows(reader), headers, (writer, row) => {
                for (int column = 0; column < headers.Length; column++) {
                    if (row.IsDBNull(column)) writer.WriteBlank();
                    else writer.Write(row.GetValue(column));
                }
            }, cancellationToken);
        } catch { _workbook.Abort(); throw; }
    }

    private static IEnumerable<IDataRecord> ReadRows(IDataReader reader) {
        while (reader.Read()) yield return reader;
    }

    /// <summary>Abandons pending output, including already completed tables. Dispose releases the writer without publishing its staged file.</summary>
    public void Abort() {
        _workbook.Abort();
        _activeTable?.Dispose();
        _activeTable = null;
    }

    /// <summary>Completes and publishes a successful workbook, or discards a failed file export.</summary>
    public void Dispose() {
        _activeTable?.Dispose();
        _activeTable = null;
        _workbook.Dispose();
    }
}
