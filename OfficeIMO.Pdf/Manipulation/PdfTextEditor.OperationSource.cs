namespace OfficeIMO.Pdf;

internal static partial class PdfTextEditor {
    /// <summary>
    /// Reuses the original artifact's readers within one synchronous edit. Rewritten bytes must
    /// receive a new reader; visual extraction and the caller's original read contract stay separate.
    /// </summary>
    private sealed class TextEditSource {
        private readonly byte[] _bytes;
        private readonly PdfLoadOptions? _readOptions;
        private PdfReadDocument? _document;
        private PdfReadDocument? _visualDocument;

        internal TextEditSource(byte[] bytes, PdfLoadOptions? readOptions) {
            _bytes = bytes;
            _readOptions = readOptions;
        }

        internal PdfReadDocument Document => _document ??= PdfReadDocument.Open(_bytes, _readOptions);

        internal PdfReadDocument VisualDocument => _visualDocument ??= OpenForVisualTextEditing(_bytes, _readOptions);
    }
}
