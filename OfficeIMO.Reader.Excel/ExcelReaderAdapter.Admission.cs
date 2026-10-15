using OfficeIMO.Excel;
using OfficeIMO.Reader.FormatInternals;

namespace OfficeIMO.Reader.Excel;

internal static partial class ExcelReaderAdapter {
    // Native row extraction is lazy: translate each pull before routing or aggregate failure policy can handle it.
    private static IEnumerable<ReaderChunk> Extract(ExcelDocumentReader reader, string sourceName,
        ReaderOptions readerOptions, ReaderExcelOptions options, IReadOnlyList<string>? legacyWarnings,
        CancellationToken cancellationToken) {
        using var iterator = ExtractNative(reader, sourceName, readerOptions, options, legacyWarnings, cancellationToken).GetEnumerator();
        while (NativeReaderAdmission.Load(iterator.MoveNext)) yield return iterator.Current;
    }
}
