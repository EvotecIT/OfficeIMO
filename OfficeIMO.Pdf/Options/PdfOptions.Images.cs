namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    private PdfImageOptimizationOptions? _imageOptimization;

    /// <summary>
    /// Optional placement-aware image optimization used while generating image XObjects.
    /// The value is snapshotted; optimization is disabled when this option is null or disabled.
    /// </summary>
    public PdfImageOptimizationOptions? ImageOptimization {
        get => _imageOptimization?.Clone();
        set => _imageOptimization = value?.Clone();
    }

    internal PdfImageOptimizationOptions? ImageOptimizationSnapshot => _imageOptimization?.Clone();
}
