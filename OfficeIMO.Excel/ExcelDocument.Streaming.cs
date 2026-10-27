using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Creates a bounded-memory XLSX writer. Dispose publishes the completed file atomically.</summary>
    public static ExcelWorkbookWriter CreateWriter(string path, ExcelWriterOptions? options = null, OfficeTheme? theme = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Supply an XLSX destination path.", nameof(path));
        string fullPath = Path.GetFullPath(path);
        var snapshot = CreateStreamingOptions(options);
        var output = OfficeFileCommit.CreateTemporaryFile(fullPath, FileOptions.SequentialScan, out string temporaryPath);
        try { return new ExcelWorkbookWriter(new StreamingWorkbook(output, snapshot, theme ?? OfficeTheme.Modern, false, fullPath, temporaryPath)); }
        catch { output.Dispose(); File.Delete(temporaryPath); throw; }
    }

    /// <summary>Creates a bounded-memory XLSX writer over a writable stream, including non-seekable streams.</summary>
    public static ExcelWorkbookWriter CreateWriter(Stream stream, ExcelWriterOptions? options = null, OfficeTheme? theme = null, bool leaveOpen = false) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(stream));
        return new ExcelWorkbookWriter(new StreamingWorkbook(stream, CreateStreamingOptions(options), theme ?? OfficeTheme.Modern, leaveOpen));
    }

    private static ExcelWriterOptions CreateStreamingOptions(ExcelWriterOptions? options) {
        var snapshot = CreateRowWriteOptions(options ?? new ExcelWriterOptions { CreateTable = true, UseSharedStrings = false });
        if (snapshot.AutoFit || snapshot.UseSharedStrings)
            throw new ArgumentException("The streaming writer requires AutoFit=false and UseSharedStrings=false.", nameof(options));
        if (snapshot.CreateTable && !snapshot.IncludeHeaders)
            throw new ArgumentException("Streaming tables require column headers.", nameof(options));
        snapshot.RequireStreaming = true;
        snapshot.Styles = snapshot.Styles?.ToDictionary(item => item.Key, item => item.Value.Snapshot(), StringComparer.Ordinal);
        snapshot.ColumnStyles = snapshot.ColumnStyles?.ToDictionary(item => item.Key, item => item.Value);
        return snapshot;
    }

    /// <summary>Owns the existing native row writer's incremental package and deferred workbook metadata.</summary>
    internal sealed class StreamingWorkbook : IDisposable {
        private readonly Stream _output;
        private readonly ExcelPositionReportingWriteStream? _position;
        private readonly ZipArchive _archive;
        private readonly ExcelWriterOptions _options;
        private readonly OfficeTheme _theme;
        private readonly bool _leaveOpen;
        private readonly string? _path;
        private readonly string? _temporaryPath;
        private readonly List<DirectDataSetSheetModel> _sheets = new();
        private readonly HashSet<string> _names = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _tableNames = new(StringComparer.OrdinalIgnoreCase);
        private bool _failed;
        private bool _disposed;
        private int _busy;

        internal StreamingWorkbook(Stream output, ExcelWriterOptions options, OfficeTheme theme, bool leaveOpen,
            string? path = null, string? temporaryPath = null) {
            _output = output; _options = options; _theme = theme; _leaveOpen = leaveOpen;
            _path = path; _temporaryPath = temporaryPath;
            if (output.CanSeek) PrepareDestinationStreamForWrite(output);
            _position = output.CanSeek ? null : new ExcelPositionReportingWriteStream(output);
            _archive = new ZipArchive(_position ?? output, ZipArchiveMode.Create, leaveOpen: true);
        }

        internal async Task WriteTable<T>(string name, IAsyncEnumerable<T> rows, IReadOnlyList<string> headers,
            Action<ExcelTabularRowWriter, T> writeRow, CancellationToken cancellationToken) {
            if (_disposed) throw new ObjectDisposedException(nameof(ExcelWorkbookWriter));
            if (Interlocked.CompareExchange(ref _busy, 1, 0) != 0)
                throw new InvalidOperationException("Complete the current table before writing another table.");
            try {
                if (_failed) throw new InvalidOperationException("A failed streaming workbook cannot accept more tables.");
                if (rows == null) throw new ArgumentNullException(nameof(rows));
                if (headers == null) throw new ArgumentNullException(nameof(headers));
                if (writeRow == null) throw new ArgumentNullException(nameof(writeRow));
                if (headers.Count == 0 || headers.Count > ExcelA1.MaxColumns)
                    throw new ArgumentException("Supply between one and 16384 column headers.", nameof(headers));
                ValidateTabularHeaders(headers);
                if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Supply a worksheet name.", nameof(name));
                cancellationToken.ThrowIfCancellationRequested();
                string sheetName = DirectDataSetWorkbookModel.SanitizeSheetName(name);
                if (!_names.Add(sheetName)) throw new ArgumentException("Worksheet names must be unique after sanitizing.", nameof(name));
                var names = headers.ToArray();
                var types = Enumerable.Repeat(typeof(object), names.Length).ToArray();
                var table = DirectDataSetTableModel.FromRows(names, types, Array.Empty<object?[]>());
                var styles = ExcelTabularStylePlan.Create(_options, names.Length, cancellationToken);
                string? tableName = _options.CreateTable
                    ? DirectDataSetWorkbookModel.SanitizeTableName(_options.TableName ?? name) : null;
                if (tableName != null && !_tableNames.Add(tableName))
                    throw new ArgumentException("Table names must be unique after sanitizing.", nameof(name));
                var sheet = new DirectDataSetSheetModel(_sheets.Count + 1, sheetName, tableName, "", table,
                    _options.TableStyle, _options.IncludeHeaders, _options.IncludeAutoFilter, _options.CreateTable,
                    false, false, null, _options.UseCellValueNumberFormats, includeCellReferences: _options.IncludeCellReferences);
                var model = DirectDataSetWorkbookModel.CreateStreaming(new[] { sheet }, _options.DateSystem, styles);
                await DirectDataSetWorkbookWriter.WriteStreamingWorksheet(_archive, model, rows, writeRow, cancellationToken).ConfigureAwait(false);
                _sheets.Add(sheet);
            } catch { _failed = true; throw; }
            finally { Volatile.Write(ref _busy, 0); }
        }

        internal void Abort() => _failed = true;

        public void Dispose() {
            if (_disposed) return;
            if (Interlocked.CompareExchange(ref _busy, 1, 0) != 0)
                throw new InvalidOperationException("Complete the current table before disposing the writer.");
            _disposed = true;
            try {
                try {
                    if (!_failed) {
                        if (_sheets.Count == 0) { _failed = true; return; }
                        var styles = ExcelTabularStylePlan.Create(_options, _sheets[0].Table.ColumnCount, CancellationToken.None);
                        var model = DirectDataSetWorkbookModel.CreateStreaming(_sheets, _options.DateSystem, styles);
                        DirectDataSetWorkbookWriter.CompleteStreamingWorkbook(_archive, model, _theme);
                    }
                } finally { _archive.Dispose(); }
                if (!_leaveOpen) _output.Dispose();
                if (!_failed && _path != null)
                    OfficeFileCommit.CommitTemporaryFileAtomically(_temporaryPath!, _path);
            } finally {
                _position?.Dispose();
                if (!_leaveOpen) _output.Dispose();
                if (_temporaryPath != null && File.Exists(_temporaryPath)) File.Delete(_temporaryPath);
            }
        }
    }
}
