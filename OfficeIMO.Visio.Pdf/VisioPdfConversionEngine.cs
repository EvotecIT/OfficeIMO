using OfficeIMO;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Visio.Pdf;

internal static partial class VisioPdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(
        VisioDocument document,
        VisioToPdfOptions? options,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        VisioToPdfOptions operation = options ?? new VisioToPdfOptions();
        operation.Validate();
        cancellationToken.ThrowIfCancellationRequested();

        if (operation.Mode == VisioPdfProjectionMode.DiagramPages) return ConvertDiagramPages(document, operation, cancellationToken);

        OfficeDocumentModel normalized = document.ToOfficeDocumentModel(
            operation.SourceName,
            operation.VisioOptions,
            cancellationToken);
        return PdfCore.OfficeDocumentModelPdfExtensions.ToPdfDocumentResult(normalized, operation.ProjectionOptions, cancellationToken);
    }
}
