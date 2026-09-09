using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageImageRenderer {
    internal static PdfPageRenderResult RenderPrintPage(Func<CancellationToken, byte[]> getPdf, int pageNumber,
        PdfPagePrintOptions? options, PdfLoadOptions readOptions, CancellationToken cancellationToken) {
        PdfPagePrintOptions effective = options ?? new();
        PdfPageRenderOptions rendering = effective.ToRenderOptions();
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(rendering.RenderTimeout, cancellationToken);
        execution.Token.ThrowIfCancellationRequested();
        PdfReadDocument document = PdfReadDocument.Open(getPdf(execution.Token), readOptions, execution.Token);
        ValidatePageNumber(document, pageNumber);
        PdfPermissionAuthorization.DemandPrinting(document.Security, readOptions.PermissionPolicy);
        if (effective.Dpi > 150 && document.Security.HasEncryption && !document.Security.HasOwnerAuthorization &&
            readOptions.PermissionPolicy != PdfPermissionPolicy.IgnoreRestrictions && document.Security.AllowsHighQualityPrinting != true) {
            throw new PdfPermissionDeniedException(PdfStandardPermissions.HighQualityPrint,
                document.Security.PasswordAuthenticationRole, "This document permits raster printing only at 150 DPI or below.");
        }
        PdfPageRenderResult result = RenderPage(document, pageNumber, rendering, execution.Token, forDisplay: true);
        execution.ThrowIfCancellationRequested();
        return result;
    }

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
