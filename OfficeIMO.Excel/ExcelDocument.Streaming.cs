#pragma warning disable RS0026 // A required format in the first position keeps every existing call unambiguous.
using System.IO.Compression;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Creates a bounded-memory XLSX writer. Dispose publishes the completed file atomically.</summary>
    public static ExcelWorkbookWriter CreateWriter(string path, ExcelWriterOptions? options = null, OfficeTheme? theme = null) =>
        CreateWriter(ExcelFileFormat.Xlsx, path, options, theme);

    /// <summary>Creates a row writer for the selected Excel format. Dispose publishes the file atomically.</summary>
    /// <remarks>XLSX and XLSB stream rows. XLS stages native BIFF records; text uses inline labels unless UseSharedStrings is set or it exceeds 255 characters.</remarks>
    public static ExcelWorkbookWriter CreateWriter(ExcelFileFormat format, string path, ExcelWriterOptions? options = null, OfficeTheme? theme = null) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        string fullPath = Path.GetFullPath(path);
        var snapshot = CreateStreamingOptions(options, format);
        if (snapshot.NoClobber && File.Exists(fullPath)) throw new IOException("The destination file already exists.");
        var output = OfficeFileCommit.CreateTemporaryFile(fullPath, FileOptions.SequentialScan, out string temporaryPath);
        try { return new ExcelWorkbookWriter(new StreamingWorkbook(output, snapshot, theme ?? OfficeTheme.Modern, false, fullPath, temporaryPath, format)); }
        catch { output.Dispose(); File.Delete(temporaryPath); throw; }
    }

    /// <summary>Creates a bounded-memory XLSX writer over a writable stream, including non-seekable streams.</summary>
    public static ExcelWorkbookWriter CreateWriter(Stream stream, ExcelWriterOptions? options = null, OfficeTheme? theme = null, bool leaveOpen = false) =>
        CreateWriter(ExcelFileFormat.Xlsx, stream, options, theme, leaveOpen);

    /// <summary>Creates a row writer for the selected format over a writable, optionally non-seekable stream.</summary>
    /// <remarks>Binary formats default to plain worksheets and reject tables, automatic sizing and declared styles.</remarks>
    public static ExcelWorkbookWriter CreateWriter(ExcelFileFormat format, Stream stream, ExcelWriterOptions? options = null, OfficeTheme? theme = null, bool leaveOpen = false) {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(stream));
        var snapshot = CreateStreamingOptions(options, format);
        if (snapshot.NoClobber) throw new ArgumentException("NoClobber requires a file destination.", nameof(options));
        return new ExcelWorkbookWriter(new StreamingWorkbook(stream, snapshot, theme ?? OfficeTheme.Modern, leaveOpen, format: format));
    }

    private static ExcelWriterOptions CreateStreamingOptions(ExcelWriterOptions? options, ExcelFileFormat format) {
        if (format is not ExcelFileFormat.Xlsx and not ExcelFileFormat.Xlsb and not ExcelFileFormat.Xls)
            throw new ArgumentOutOfRangeException(nameof(format));
        var snapshot = CreateRowWriteOptions(options ?? new ExcelWriterOptions { CreateTable = format == ExcelFileFormat.Xlsx, UseSharedStrings = false });
        if (snapshot.UseSharedStrings && format != ExcelFileFormat.Xls)
            throw new ArgumentException("The streaming writer requires UseSharedStrings=false.", nameof(options));
        snapshot.Title = string.IsNullOrWhiteSpace(snapshot.Title) ? null : snapshot.Title;
        var (firstRow, firstColumn) = ExcelA1.ParseCellRef(snapshot.Address);
        if (firstRow < 1 || firstRow > ExcelA1.MaxRows || firstColumn < 1 || firstColumn > ExcelA1.MaxColumns)
            throw new ArgumentOutOfRangeException(nameof(options.Address));
        if (!snapshot.IncludeCellReferences && firstColumn != 1)
            throw new ArgumentException("Column offsets require explicit cell references.", nameof(options));
        if (snapshot.ColumnFormat != null) {
            var formats = new ExcelColumnFormatPlan();
            foreach (ExcelColumnFormatRule rule in snapshot.ColumnFormat.Rules)
                formats.AddFormat(rule.Header, rule.ResolveNumberFormat(), rule.IncludeHeader, rule.AutoFit);
            snapshot.ColumnFormat = formats;
        }
        snapshot.DocumentProperties = snapshot.DocumentProperties?.ToDictionary(item => item.Key, item => item.Value, StringComparer.OrdinalIgnoreCase);
        ValidateStreamingDocumentProperties(snapshot.DocumentProperties);
        if (snapshot.SaveOptions?.MaxInMemoryPackageBytes <= 0 || snapshot.SaveOptions?.MaxTemporaryPackageBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "Save resource limits must be positive or null.");
        if (format != ExcelFileFormat.Xlsx && (snapshot.CreateTable || snapshot.AutoFit || snapshot.BoldHeaders ||
            snapshot.FreezeHeader || snapshot.FreezeFirstColumn ||
            snapshot.Title != null || firstRow != 1 || firstColumn != 1 || snapshot.ColumnFormat != null ||
            snapshot.DocumentProperties?.Count > 0 || snapshot.SaveOptions != null || snapshot.Styles?.Count > 0 ||
            snapshot.DefaultRowStyle != null || snapshot.ColumnStyles?.Count > 0))
            throw new NotSupportedException("Binary row writers support plain worksheets at A1 without tables, sizing, titles, freeze panes, declared or column styles, document properties or final save policies.");
        if (format == ExcelFileFormat.Xls && snapshot.RequireStreaming)
            throw new NotSupportedException("XLS requires native record buffering. Use XLSX or XLSB for bounded row streaming.");
        snapshot.RequireStreaming = format != ExcelFileFormat.Xls;
        snapshot.Styles = snapshot.Styles?.ToDictionary(item => item.Key, item => item.Value.Snapshot(), StringComparer.Ordinal);
        snapshot.ColumnStyles = snapshot.ColumnStyles?.ToDictionary(item => item.Key, item => item.Value);
        return snapshot;
    }

    /// <summary>Owns the existing native row writer's incremental package and deferred workbook metadata.</summary>
    internal sealed class StreamingWorkbook : IDisposable {
        private readonly Stream _output;
        private readonly Stream? _stagedDestination;
        private readonly string? _stagedPath;
        private readonly ExcelPositionReportingWriteStream? _position;
        private ZipArchive? _archive;
        private readonly ExcelFileFormat _format;
        private BinaryStreamingWorkbook? _binary;
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
        private ExcelTabularStylePlan? _styles;

        internal long? MaximumTemporaryBytes {
            get {
                if (_disposed) throw new ObjectDisposedException(nameof(ExcelWorkbookWriter));
                if (_failed) throw new InvalidOperationException("A failed streaming workbook cannot accept more tables.");
                if (Volatile.Read(ref _busy) != 0) throw new InvalidOperationException("Complete the current table before beginning another table.");
                return _options.SaveOptions == null ? ExcelSaveOptions.DefaultMaxTemporaryPackageBytes : _options.SaveOptions.MaxTemporaryPackageBytes;
            }
        }

        internal StreamingWorkbook(Stream output, ExcelWriterOptions options, OfficeTheme theme, bool leaveOpen,
            string? path = null, string? temporaryPath = null, ExcelFileFormat format = ExcelFileFormat.Xlsx) {
            _options = options; _theme = theme; _leaveOpen = leaveOpen;
            _path = path; _temporaryPath = temporaryPath; _format = format;
            if (path == null && NeedsStreamingFinalization(options)) {
                _stagedDestination = output;
                _stagedPath = Path.Combine(Path.GetTempPath(), "OfficeIMO-Excel-Package-" + Guid.NewGuid().ToString("N") + ".xlsx");
                _output = new FileStream(_stagedPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
            } else _output = output;
            _position = _output.CanSeek ? null : new ExcelPositionReportingWriteStream(_output);
        }

        internal async Task<ExcelDataSetImportResult> WriteTable<T>(string name, IAsyncEnumerable<T> rows, IReadOnlyList<string> headers,
            Action<ExcelTabularRowWriter, T> writeRow, CancellationToken cancellationToken) {
            if (_disposed) throw new ObjectDisposedException(nameof(ExcelWorkbookWriter));
            if (Interlocked.CompareExchange(ref _busy, 1, 0) != 0)
                throw new InvalidOperationException("Complete the current table before writing another table.");
            try {
                if (_failed) throw new InvalidOperationException("A failed streaming workbook cannot accept more tables.");
                ArgumentNullException.ThrowIfNull(rows);
                ArgumentNullException.ThrowIfNull(headers);
                ArgumentNullException.ThrowIfNull(writeRow);
                if (headers.Count == 0 || headers.Count > (_format == ExcelFileFormat.Xls ? 256 : ExcelA1.MaxColumns))
                    throw new ArgumentException(_format == ExcelFileFormat.Xls ? "Supply between one and 256 column headers." : "Supply between one and 16384 column headers.", nameof(headers));
                ValidateTabularHeaders(headers);
                ArgumentException.ThrowIfNullOrWhiteSpace(name);
                cancellationToken.ThrowIfCancellationRequested();
                string sheetName = DirectDataSetWorkbookModel.SanitizeSheetName(name);
                if (!_names.Add(sheetName)) throw new ArgumentException("Worksheet names must be unique after sanitizing.", nameof(name));
                var names = headers.ToArray();
                var types = Enumerable.Repeat(typeof(object), names.Length).ToArray();
                var table = DirectDataSetTableModel.FromRows(names, types, Array.Empty<object?[]>());
                var styleOptions = PrepareStreamingStyles(_options, names);
                var styles = ExcelTabularStylePlan.Create(styleOptions, names.Length, cancellationToken, _theme);
                _styles ??= styles;
                var layout = new StreamingWorksheetLayout(_options, names, styles);
                string? tableName = _options.CreateTable
                    ? DirectDataSetWorkbookModel.SanitizeTableName(string.IsNullOrWhiteSpace(_options.TableName) ? name : _options.TableName!) : null;
                if (tableName != null) {
                    if (string.IsNullOrWhiteSpace(_options.TableName))
                        tableName = DirectDataSetWorkbookModel.GetUniqueName(tableName, _tableNames, 255);
                    else if (!_tableNames.Add(tableName))
                        throw new ArgumentException("Table names must be unique after sanitizing.", nameof(name));
                }
                var sheet = new DirectDataSetSheetModel(_sheets.Count + 1, sheetName, tableName, "", table,
                    _options.TableStyle, _options.IncludeHeaders, _options.IncludeAutoFilter, _options.CreateTable,
                    false, false, null, _options.UseCellValueNumberFormats,
                    showFirstColumn: _options.ShowFirstColumn, showLastColumn: _options.ShowLastColumn,
                    showRowStripes: _options.ShowRowStripes, showColumnStripes: _options.ShowColumnStripes,
                    includeCellReferences: _options.IncludeCellReferences) { StreamingLayout = layout };
                var model = DirectDataSetWorkbookModel.CreateStreaming(new[] { sheet }, _options.DateSystem, styles);
                if (_archive == null && _format != ExcelFileFormat.Xls) {
                    if (_output.CanSeek) PrepareDestinationStreamForWrite(_output);
                    _archive = new ZipArchive(_position ?? _output, ZipArchiveMode.Create, leaveOpen: true);
                }
                int rowCount;
                if (_format == ExcelFileFormat.Xlsx) {
                    rowCount = await DirectDataSetWorkbookWriter.WriteStreamingWorksheet(_archive!, model, rows, writeRow,
                        layout.AutoFitColumns.Any(enabled => enabled), _theme, cancellationToken).ConfigureAwait(false);
                } else {
                    _binary ??= new BinaryStreamingWorkbook(_format, _archive, _options, BinaryStreamingStyles.Get(
                        DirectDataSetWorkbookWriter.CreateNativeStreamingStylesXml(model, _theme),
                        DirectDataSetWorkbookWriter.CreateNativeStreamingTheme(_theme), _theme));
                    rowCount = await _binary.WriteTable(sheetName, names, rows, writeRow, cancellationToken).ConfigureAwait(false);
                }
                _sheets.Add(sheet);
                return new ExcelDataSetImportResult(sheetName, tableName,
                    layout.Range(names.Length, rowCount, _options.IncludeHeaders), rowCount, names.Length);
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
                    if (!_failed && _sheets.Count == 0)
                        throw new InvalidOperationException("Write at least one table before disposing the writer; a workbook needs a worksheet.");
                    if (!_failed) {
                        var model = DirectDataSetWorkbookModel.CreateStreaming(_sheets, _options.DateSystem, _styles);
                        if (_format == ExcelFileFormat.Xlsx) DirectDataSetWorkbookWriter.CompleteStreamingWorkbook(_archive!, model, _theme);
                        else _binary!.Complete(_output, _sheets.Select(sheet => sheet.SheetName).ToArray());
                    }
                } finally { try { _archive?.Dispose(); } finally { _binary?.Dispose(); } }
                if (!_leaveOpen || _stagedPath != null) _output.Dispose();
                if (!_failed && _format == ExcelFileFormat.Xlsx && (_temporaryPath ?? _stagedPath) is string packagePath)
                    FinalizeStreamingFile(packagePath, _options);
                if (!_failed && _stagedDestination != null) {
                    if (_stagedDestination.CanSeek) PrepareDestinationStreamForWrite(_stagedDestination);
                    using var completed = File.OpenRead(_stagedPath!);
                    completed.CopyTo(_stagedDestination);
                }
                if (!_failed && _path != null)
                    OfficeFileCommit.CommitTemporaryFileAtomically(_temporaryPath!, _path,
                        _options.NoClobber ? OfficeFileCommit.ConflictPolicy.FailIfExists : OfficeFileCommit.ConflictPolicy.Replace);
            } finally {
                _position?.Dispose();
                if (!_leaveOpen || _stagedPath != null) _output.Dispose();
                if (!_leaveOpen) _stagedDestination?.Dispose();
                if (_stagedPath != null && File.Exists(_stagedPath)) File.Delete(_stagedPath);
                if (_temporaryPath != null && File.Exists(_temporaryPath)) File.Delete(_temporaryPath);
            }
        }
    }
}

#pragma warning restore RS0026
