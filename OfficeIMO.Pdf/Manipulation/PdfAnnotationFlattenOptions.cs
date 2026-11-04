namespace OfficeIMO.Pdf;

/// <summary>Selects visual annotations to paint into page content and remove as live annotations.</summary>
public sealed class PdfAnnotationFlattenOptions : IPdfOperationCancellation {
    // Operation cancellation; public methods take it as a parameter and bind it to a per-call copy.
    System.Threading.CancellationToken IPdfOperationCancellation.CancellationToken { get => CancellationToken; set => CancellationToken = value; }

    internal System.Threading.CancellationToken CancellationToken { get; set; }

    object IPdfOperationCancellation.CloneForOperation() => MemberwiseClone();

    /// <summary>Optional exact indirect annotation object number.</summary>
    public int? ObjectNumber { get; set; }
    /// <summary>Optional one-based page number.</summary>
    public int? PageNumber { get; set; }
    /// <summary>Optional annotation subtype.</summary>
    public string? Subtype { get; set; }
}
