using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>
/// Configures page size normalization for existing PDF pages.
/// </summary>
public sealed class PdfPageResizeOptions : IPdfOperationCancellation {
    /// <summary>
    /// Creates resize options for the supplied target page size.
    /// </summary>
    /// <param name="pageSize">Target page size in PDF points.</param>
    public PdfPageResizeOptions(PdfPageSize pageSize) {
        Guard.Positive(pageSize.Width.ToPoints(), nameof(pageSize));
        Guard.Positive(pageSize.Height.ToPoints(), nameof(pageSize));
        PageSize = pageSize;
    }

    // Operation cancellation; public methods take it as a parameter and bind it to a per-call copy.
    System.Threading.CancellationToken IPdfOperationCancellation.CancellationToken { get => CancellationToken; set => CancellationToken = value; }

    internal System.Threading.CancellationToken CancellationToken { get; set; }

    object IPdfOperationCancellation.CloneForOperation() => MemberwiseClone();

    /// <summary>Target page size in PDF points.</summary>
    public PdfPageSize PageSize { get; set; }

    /// <summary>Margin in PDF points kept around fitted content.</summary>
    public OfficeLength Margin { get; set; }

    /// <summary>Scaling behavior used for the original page content.</summary>
    public PdfPageResizeMode Mode { get; set; } = PdfPageResizeMode.Fit;
}
