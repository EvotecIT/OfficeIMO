namespace OfficeIMO.Pdf;

internal readonly struct PdfPageShadingPatternResource {
    public PdfPageShadingPatternResource(PdfPageShadingResource shading, Matrix2D matrix, bool hasExactMatrix = true) {
        Shading = shading;
        Matrix = matrix;
        IsSupported = true;
        HasExactMatrix = hasExactMatrix;
    }

    public static PdfPageShadingPatternResource Unsupported => default;

    public PdfPageShadingResource Shading { get; }

    public Matrix2D Matrix { get; }

    public bool IsSupported { get; }

    public bool HasExactMatrix { get; }

    public bool SupportsExactType3Projection => IsSupported && HasExactMatrix && Shading.SupportsExactType3Projection;
}
