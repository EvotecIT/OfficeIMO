namespace OfficeIMO.Pdf;

/// <summary>
/// Configures optional behavior for first-party PDF page import operations.
/// </summary>
public sealed class PdfPageImportOptions : IPdfOperationCancellation {
    // Operation cancellation; public methods take it as a parameter and bind it to a per-call copy.
    System.Threading.CancellationToken IPdfOperationCancellation.CancellationToken { get => CancellationToken; set => CancellationToken = value; }

    internal System.Threading.CancellationToken CancellationToken { get; set; }

    object IPdfOperationCancellation.CloneForOperation() => MemberwiseClone();

    /// <summary>
    /// Gets or sets read options used for the imported source PDF, including its password and permission policy.
    /// Target-document credentials remain controlled by the document or the corresponding <c>Try*</c> operation.
    /// </summary>
    public PdfLoadOptions? SourceReadOptions { get; set; }

    /// <summary>
    /// Gets or sets whether supported visual annotations from imported source pages should be painted into page content.
    /// Existing target-document annotations are not flattened by this option.
    /// </summary>
    public bool FlattenVisualAnnotations { get; set; }
}
