using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageImageRenderer {
    internal static PdfPageRenderResult RenderDisplayPage(Func<CancellationToken, byte[]> getPdf, int pageNumber,
        PdfPageDisplayOptions? options, PdfLoadOptions readOptions, CancellationToken cancellationToken) {
        PdfPageRenderOptions rendering = (options ?? new PdfPageDisplayOptions()).ToRenderOptions();
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(rendering.RenderTimeout, cancellationToken);
        try {
            execution.Token.ThrowIfCancellationRequested();
            PdfReadDocument document = PdfReadDocument.Open(getPdf(execution.Token), readOptions, execution.Token);
            ValidatePageNumber(document, pageNumber);
            PdfPageRenderResult result = RenderPage(document, pageNumber, rendering, execution.Token, forDisplay: true);
            execution.ThrowIfCancellationRequested();
            return result;
        } catch (OperationCanceledException error) when (execution.IsTimeoutCancellation(error)) {
            throw execution.CreateTimeoutException(error);
        }
    }
}
