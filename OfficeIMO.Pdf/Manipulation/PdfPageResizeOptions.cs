namespace OfficeIMO.Pdf;

/// <summary>
/// Configures page size normalization for existing PDF pages.
/// </summary>
public sealed class PdfPageResizeOptions {
    /// <summary>
    /// Creates resize options for the supplied target page size.
    /// </summary>
    /// <param name="pageSize">Target page size in PDF points.</param>
    public PdfPageResizeOptions(PageSize pageSize) {
        Guard.Positive(pageSize.Width, nameof(pageSize));
        Guard.Positive(pageSize.Height, nameof(pageSize));
        PageSize = pageSize;
    }

    /// <summary>Target page size in PDF points.</summary>
    public PageSize PageSize { get; set; }

    /// <summary>Margin in PDF points kept around fitted content.</summary>
    public double Margin { get; set; }

    /// <summary>Scaling behavior used for the original page content.</summary>
    public PdfPageResizeMode Mode { get; set; } = PdfPageResizeMode.Fit;
}
