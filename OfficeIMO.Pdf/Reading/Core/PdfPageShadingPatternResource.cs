namespace OfficeIMO.Pdf;

internal readonly struct PdfPageShadingPatternResource {
    public PdfPageShadingPatternResource(PdfPageShadingResource shading, Matrix2D matrix) {
        Shading = shading;
        Matrix = matrix;
    }

    public PdfPageShadingResource Shading { get; }

    public Matrix2D Matrix { get; }
}
