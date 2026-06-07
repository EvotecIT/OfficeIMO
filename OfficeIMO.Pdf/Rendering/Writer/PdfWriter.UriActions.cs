namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static void ValidateUriActionLinks(IReadOnlyList<LayoutResult.Page> pages, PdfOptions options) {
        bool hasCatalogUriBase = !string.IsNullOrWhiteSpace(options.CatalogUriBaseSnapshot);
        foreach (var page in pages) {
            foreach (var annotation in page.Annotations) {
                if (string.IsNullOrWhiteSpace(annotation.Uri)) {
                    continue;
                }

                Guard.UriAction(annotation.Uri, nameof(annotation.Uri));
                if (!Uri.TryCreate(annotation.Uri, UriKind.Absolute, out _) && !hasCatalogUriBase) {
                    throw new ArgumentException("Relative PDF URI link targets require PdfOptions.CatalogUriBase.");
                }
            }
        }
    }
}
