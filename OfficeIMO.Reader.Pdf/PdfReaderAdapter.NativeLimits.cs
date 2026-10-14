using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    // A native quota must remain terminal even through folder/batch recovery policies.
    private static ReaderAdmissionLimitException TranslateNativeLimit(PdfReadLimitException error) =>
        new ReaderAdmissionLimitException("Pdf." + error.Kind, error.Limit, error.Actual, error);

    private static T WithNativeLimits<T>(Func<T> operation) {
        try { return operation(); }
        catch (PdfReadLimitException error) { throw TranslateNativeLimit(error); }
    }

    private static IEnumerable<T> WithNativeLimits<T>(IEnumerable<T> source) {
        using IEnumerator<T> iterator = source.GetEnumerator();
        while (true) {
            bool more;
            try { more = iterator.MoveNext(); }
            catch (PdfReadLimitException error) { throw TranslateNativeLimit(error); }
            if (!more) yield break;
            yield return iterator.Current;
        }
    }

    private static PdfDocument OpenReaderPdf(string path, ReaderOptions options, ReaderPdfOptions pdfOptions,
        CancellationToken token) => WithNativeLimits(() => PdfDocument.Load(path, CreatePdfLoadOptions(options, pdfOptions), token));

}
