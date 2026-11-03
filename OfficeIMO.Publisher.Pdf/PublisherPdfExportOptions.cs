using OfficeIMO.Pdf;

namespace OfficeIMO.Publisher.Pdf;

/// <summary>Controls page-preserving Publisher-to-PDF conversion.</summary>
public sealed class PublisherPdfExportOptions {
    /// <summary>PDF generation settings. Source page sizes and zero margins take precedence.</summary>
    public PdfCreateOptions? PdfOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates an independent copy of these settings.</summary>
    public PublisherPdfExportOptions Clone() => new PublisherPdfExportOptions {
        PdfOptions = PdfOptions?.Clone(),
        FailOnLoss = FailOnLoss
    };
}
