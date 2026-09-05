using OfficeIMO;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Visio.Pdf;

internal static class VisioPdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(
        VisioDocument document,
        VisioToPdfOptions? options,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        VisioToPdfOptions operation = options ?? new VisioToPdfOptions();
        operation.Validate();
        cancellationToken.ThrowIfCancellationRequested();

        OfficeDocumentModel normalized = document.ToOfficeDocumentModel(
            operation.SourceName,
            operation.VisioOptions,
            cancellationToken);
        return PdfCore.OfficeDocumentModelPdfExtensions.ToPdfDocumentResult(normalized, operation.ProjectionOptions);
    }
}
