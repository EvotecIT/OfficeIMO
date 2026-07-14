namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal (double Width, double Height, Matrix2D Transform) GetImportGeometry() {
        (double Width, double Height) size = GetVisualPageSize();
        return (size.Width, size.Height, GetVisualPageTransform());
    }
}
