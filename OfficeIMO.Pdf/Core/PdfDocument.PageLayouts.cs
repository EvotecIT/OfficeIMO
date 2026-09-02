using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Reads page geometry without extracting document content or requiring copy permission.
    /// </summary>
    public IReadOnlyList<PdfPageLayoutInfo> GetPageLayouts(PdfLoadOptions? options = null) =>
        GetPageLayouts(options, CancellationToken.None);

    /// <summary>
    /// Reads page geometry with cooperative cancellation, without extracting document content or
    /// requiring copy permission.
    /// </summary>
    public IReadOnlyList<PdfPageLayoutInfo> GetPageLayouts(
        PdfLoadOptions? options,
        CancellationToken cancellationToken) {
        PdfReadDocument document = GetReadDocument(options ?? ReadOptions, cancellationToken);
        return CreatePageLayouts(document, cancellationToken);
    }

    /// <summary>
    /// Reads page geometry for printing after enforcing the authenticated document permissions.
    /// </summary>
    public IReadOnlyList<PdfPageLayoutInfo> GetPrintablePageLayouts(PdfLoadOptions? options = null) =>
        GetPrintablePageLayouts(options, CancellationToken.None);

    /// <summary>
    /// Reads page geometry for printing with cooperative cancellation after enforcing the
    /// authenticated document permissions.
    /// </summary>
    public IReadOnlyList<PdfPageLayoutInfo> GetPrintablePageLayouts(
        PdfLoadOptions? options,
        CancellationToken cancellationToken) {
        PdfLoadOptions effective = PdfLoadOptions.Resolve(options ?? ReadOptions);
        PdfReadDocument document = GetReadDocument(effective, cancellationToken);
        PdfPermissionAuthorization.DemandPrinting(document.Security, effective.PermissionPolicy);
        return CreatePageLayouts(document, cancellationToken);
    }

    private static PdfPageLayoutInfo[] CreatePageLayouts(
        PdfReadDocument document,
        CancellationToken cancellationToken) {
        var layouts = new PdfPageLayoutInfo[document.Pages.Count];
        for (int index = 0; index < document.Pages.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfReadPage page = document.Pages[index];
            PdfPageGeometry geometry = page.GetGeometry();
            (double width, double height) = page.GetPageSize();
            (double visualWidth, double visualHeight) = page.GetVisualPageSize();
            layouts[index] = new PdfPageLayoutInfo(
                index + 1,
                width,
                height,
                visualWidth,
                visualHeight,
                page.GetRotationDegrees(),
                geometry.UserUnit is double userUnit && userUnit > 0D &&
                    !double.IsNaN(userUnit) && !double.IsInfinity(userUnit)
                    ? userUnit
                    : 1D,
                geometry);
        }
        return layouts;
    }
}
