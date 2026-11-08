using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Shares the public row callback across the native binary record writers.</summary>
    private sealed class BinaryStreamingWorkbook : IDisposable {
        private readonly ExcelFileFormat _format;
        private readonly ZipArchive? _archive;
        private readonly ExcelWriterOptions _options;
        private readonly Stylesheet _styles;
        private readonly string _themeXml;
        private readonly LegacyXlsWriter.StreamingWorkbook? _legacy;
        private int _sheetCount;

        internal BinaryStreamingWorkbook(ExcelFileFormat format, ZipArchive? archive, ExcelWriterOptions options,
            Stylesheet styles, string themeXml, OfficeTheme theme) {
            _format = format; _archive = archive; _options = options; _styles = styles; _themeXml = themeXml;
            if (format == ExcelFileFormat.Xls) _legacy = new LegacyXlsWriter.StreamingWorkbook(options.DateSystem, styles, theme);
        }

        internal async Task<int> WriteTable<T>(string name, string[] headers, IAsyncEnumerable<T> rows,
            Action<ExcelTabularRowWriter, T> writeRow, CancellationToken ct) {
            Stream? output = null;
            ExcelTabularRowSink? sink = null;
            bool completed = false;
            try {
                if (_legacy != null) sink = _legacy.AddSheet(name, headers.Length);
                else {
                    var entry = _archive!.CreateEntry("xl/worksheets/sheet" + (++_sheetCount).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".bin", CompressionLevel.Fastest);
                    entry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
                    output = entry.Open();
                    sink = new XlsbWorksheetPartWriter.StreamingRowSink(output, headers.Length);
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
                completed = true;
                return count;
            } finally {
                try {
                    if (_legacy == null) sink?.Dispose();
                } catch when (!completed) {
                    // A failed row remains the primary error if flushing its unfinished bytes also fails.
                } finally {
                    try { output?.Dispose(); }
                    catch when (!completed) { }
                }
            }
        }

        internal void Complete(Stream output, IReadOnlyList<string> names) {
            if (_legacy != null) _legacy.Complete(output);
            else XlsbNewPackageWriter.CompleteStreamingWorkbook(_archive!, names, _options.DateSystem, _styles, _themeXml);
        }
        public void Dispose() => _legacy?.Dispose();
    }
}
