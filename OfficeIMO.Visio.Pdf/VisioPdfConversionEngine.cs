using OfficeIMO.Reader;
using OfficeIMO.Reader.Visio;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Visio.Pdf;

internal static class VisioPdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(
        VisioDocument document,
        VisioPdfSaveOptions? options,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        VisioPdfSaveOptions operation = options ?? new VisioPdfSaveOptions();
        operation.Validate();
        cancellationToken.ThrowIfCancellationRequested();

        OfficeDocumentReadResult normalized = document.ToOfficeDocumentReadResult(
            operation.SourceName,
            operation.ReaderOptions,
            operation.VisioOptions,
            cancellationToken);
        return normalized.ToPdfDocumentResult(operation.ProjectionOptions);
    }
}
