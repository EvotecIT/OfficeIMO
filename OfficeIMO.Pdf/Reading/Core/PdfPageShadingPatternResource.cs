namespace OfficeIMO.Pdf;

internal readonly struct PdfPageShadingPatternResource {
    public PdfPageShadingPatternResource(PdfPageShadingResource shading, Matrix2D matrix) {
        Shading = shading;
        Matrix = matrix;
        IsSupported = true;
    }

    public static PdfPageShadingPatternResource Unsupported => default;

    public PdfPageShadingResource Shading { get; }

    public Matrix2D Matrix { get; }

    public bool IsSupported { get; }

    public bool SupportsExactType3Projection => IsSupported && Shading.SupportsExactType3Projection;
}
