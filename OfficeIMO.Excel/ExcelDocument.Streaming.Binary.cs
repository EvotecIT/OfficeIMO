using System.IO.Compression;
using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Shares the public row callback across the native binary record writers.</summary>
    private sealed class BinaryStreamingWorkbook : IDisposable {
        private readonly ExcelFileFormat _format;
        private readonly ZipArchive? _archive;
        private readonly ExcelWriterOptions _options;
        private readonly BinaryStreamingStyles _styles;
        private readonly LegacyXlsWriter.StreamingWorkbook? _legacy;
        private int _sheetCount;

        internal BinaryStreamingWorkbook(ExcelFileFormat format, ZipArchive? archive, ExcelWriterOptions options, BinaryStreamingStyles styles) {
            _format = format; _archive = archive; _options = options; _styles = styles;
            if (format == ExcelFileFormat.Xls) _legacy = new LegacyXlsWriter.StreamingWorkbook(options.DateSystem, styles.Xls, options.UseSharedStrings);
        }

        internal async Task<int> WriteTable<T>(string name, string[] headers, IAsyncEnumerable<T> rows,
            Action<ExcelTabularRowWriter, T> writeRow, CancellationToken ct) {
            Stream? output = null;
            StreamingWorksheetCompression? compression = null;
            ExcelTabularRowSink? sink = null;
            bool completed = false;
            try {
                if (_legacy != null) sink = _legacy.AddSheet(name, headers.Length);
                else {
                    var entry = _archive!.CreateEntry("xl/worksheets/sheet" + (++_sheetCount).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".bin", CompressionLevel.Fastest);
                    entry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
                    output = entry.Open();
                    Stream records = output;
                    if (ExcelExecutionPolicy.SupportsBackgroundProcessing) {
                        compression = new StreamingWorksheetCompression(output);
                        records = compression;
                    }
                    sink = new XlsbWorksheetPartWriter.StreamingRowSink(records, headers.Length);
                }
                if (_options.IncludeHeaders) {
                    sink.BeginRow(1);
                    for (int column = 0; column < headers.Length; column++) {
                        sink.BeginCell(column, 0U);
                        sink.WriteText(headers[column]);
                    }
                    sink.EndRow();
                }
                var writer = ExcelTabularRowWriter.Create(sink, headers.Length, _options.IncludeHeaders,
                    _options.UseCellValueNumberFormats, _options.DateSystem);
                int maximumRows = (_format == ExcelFileFormat.Xls ? 65_536 : ExcelA1.MaxRows) - (_options.IncludeHeaders ? 1 : 0);
                int count = await WriteStreamingRowCallbacks(rows, writer, writeRow, maximumRows, ct).ConfigureAwait(false);
                sink.Complete();
                compression?.Complete();
                completed = true;
                return count;
            } finally {
                try {
                    if (_legacy == null) sink?.Dispose();
                } catch when (!completed) {
                    // A failed row remains the primary error if flushing its unfinished bytes also fails.
                } finally {
                    try { compression?.Dispose(); }
                    finally {
                        try { output?.Dispose(); }
                        catch when (!completed) { }
                    }
                }
            }
        }

        internal void Complete(Stream output, IReadOnlyList<string> names) {
            if (_legacy != null) _legacy.Complete(output);
            else XlsbNewPackageWriter.CompleteStreamingWorkbook(_archive!, names, _options.DateSystem, _styles.XlsbStylesPart, _styles.ThemeXml);
        }
        public void Dispose() => _legacy?.Dispose();
    }

    /// <summary>Reuses native style parts converted from identical generated styles and theme XML.</summary>
    /// <remarks>Binary row writers reject declared styles, so ordinary writes repeat one small set of inputs.
    /// The cached parts are immutable after creation and are only read by later writes.</remarks>
    private sealed class BinaryStreamingStyles {
        private const int Capacity = 8;
        private static readonly List<BinaryStreamingStyles> Cache = new();
        private readonly string _stylesXml;
        private readonly OfficeTheme _theme;
        private readonly object _gate = new();
        private byte[]? _xlsbStylesPart;
        private LegacyXlsWriter.StreamingStyles? _xls;

        private BinaryStreamingStyles(string stylesXml, string themeXml, OfficeTheme theme) {
            _stylesXml = stylesXml; ThemeXml = themeXml; _theme = theme;
        }

        internal string ThemeXml { get; }

        internal byte[] XlsbStylesPart {
            get {
                lock (_gate) return _xlsbStylesPart ??= XlsbStylesheetPartWriter.CreateForNewWorkbook(
                    new DocumentFormat.OpenXml.Spreadsheet.Stylesheet(_stylesXml), out _);
            }
        }

        internal LegacyXlsWriter.StreamingStyles Xls {
            get {
                lock (_gate) return _xls ??= LegacyXlsWriter.StreamingStyles.Create(_stylesXml, _theme);
            }
        }

        internal static BinaryStreamingStyles Get(string stylesXml, string themeXml, OfficeTheme theme) {
            lock (Cache) {
                for (int index = 0; index < Cache.Count; index++) {
                    BinaryStreamingStyles cached = Cache[index];
                    if (string.Equals(cached._stylesXml, stylesXml, StringComparison.Ordinal)
                        && string.Equals(cached.ThemeXml, themeXml, StringComparison.Ordinal)) return cached;
                }
                var created = new BinaryStreamingStyles(stylesXml, themeXml, theme);
                if (Cache.Count == Capacity) Cache.RemoveAt(0);
                Cache.Add(created);
                return created;
            }
        }
    }
}
