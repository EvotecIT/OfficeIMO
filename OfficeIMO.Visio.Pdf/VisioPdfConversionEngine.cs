using OfficeIMO;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Visio.Pdf;

internal static partial class VisioPdfConversionEngine {
    internal static PdfCore.PdfDocumentConversionResult Convert(
        VisioDocument document,
        VisioPdfExportOptions? options,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        VisioPdfExportOptions operation = options ?? new VisioPdfExportOptions();
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
