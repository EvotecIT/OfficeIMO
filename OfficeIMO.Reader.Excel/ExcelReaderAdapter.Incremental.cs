using OfficeIMO.Reader.FormatInternals;
using OfficeIMO.Excel;

namespace OfficeIMO.Reader.Excel;

internal static partial class ExcelReaderAdapter {
    internal static IEnumerable<ReaderChunk> ReadIncremental(string path, ReaderOptions readerOptions,
        ReaderExcelOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        using ExcelDocument document = Load(path, readerOptions);
        using ExcelDocumentReader reader = NativeReaderAdmission.Load(() => document.CreateReader(options.ReadOptions));
        options = ResolveSheet(options, reader);
        foreach (var chunk in Extract(reader, path, readerOptions, options, BuildLegacyWarnings(document), cancellationToken))
            yield return chunk;
    }
}
