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
        if (document.Security.HasEncryption && !document.Security.HasOwnerAuthorization &&
            readOptions.PermissionPolicy != PdfPermissionPolicy.IgnoreRestrictions && document.Security.AllowsHighQualityPrinting != true) {
            if (effective.Dpi > 150) throw new PdfPermissionDeniedException(PdfStandardPermissions.HighQualityPrint,
                document.Security.PasswordAuthenticationRole, "This document permits raster printing only at 150 DPI or below.");
            // Enlarging a low-quality print must not recover higher-resolution source detail.
            rendering.Scale = Math.Min(rendering.Scale, 150D / 72);
        }
        PdfPageRenderResult result = RenderPage(document, pageNumber, rendering, execution.Token, forDisplay: true, printing: true);
        execution.ThrowIfCancellationRequested();
        return result;
    }

    /// <summary>
    /// Renders a display page to pixels for a caller that uses them in process, skipping the PNG encoding and decoding that
    /// <see cref="RenderDisplayPage"/> would add. The pixels match a decoded display PNG; the result carries the diagnostics.
    /// </summary>
    internal static (OfficeRasterImage Raster, PdfPageRenderResult Result) RenderDisplayRaster(Func<CancellationToken, PdfReadDocument> getDocument,
        int pageNumber, PdfPageDisplayOptions? options, CancellationToken cancellationToken) {
        OfficeRasterImage? raster = null;
        PdfPageRenderResult result = RenderDisplayPage(getDocument, pageNumber, options, cancellationToken, image => raster = image);
        return (raster ?? throw new InvalidOperationException("Page " + pageNumber + " could not be rendered: " + string.Join("; ", result.Diagnostics)), result);
    }

    internal static PdfPageRenderResult RenderDisplayPage(Func<CancellationToken, PdfReadDocument> getDocument, int pageNumber,
        PdfPageDisplayOptions? options, CancellationToken cancellationToken, Action<OfficeRasterImage>? rasterSink = null) {
        PdfPageRenderOptions rendering = (options ?? new PdfPageDisplayOptions()).ToRenderOptions();
        using OfficeImageExportExecutionScope execution = OfficeImageExportExecutionScope.Start(rendering.RenderTimeout, cancellationToken);
        try {
            execution.Token.ThrowIfCancellationRequested();
            PdfReadDocument document = getDocument(execution.Token);
            ValidatePageNumber(document, pageNumber);
            PdfPageRenderResult result = RenderPage(document, pageNumber, rendering, execution.Token, forDisplay: true, rasterSink: rasterSink);
            execution.ThrowIfCancellationRequested();
            return result;
        } catch (OperationCanceledException error) when (execution.IsTimeoutCancellation(error)) {
            throw execution.CreateTimeoutException(error);
        }
    }
}
