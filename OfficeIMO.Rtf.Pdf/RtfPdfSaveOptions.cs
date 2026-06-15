using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

/// <summary>
/// Controls conversion from an RTF document model to a first-party PDF document.
/// </summary>
public sealed class RtfPdfSaveOptions {
    /// <summary>Optional PDF engine options. The converter clones the instance before applying RTF page setup.</summary>
    public PdfCore.PdfOptions? PdfOptions { get; set; }

    /// <summary>When true, RTF hidden runs are included in PDF output. Hidden text is skipped by default.</summary>
    public bool IncludeHiddenText { get; set; }

    /// <summary>When true, supported top-level and inline PNG/JPEG images are emitted to the PDF.</summary>
    public bool IncludeImages { get; set; } = true;

    /// <summary>Default image width in PDF points when the RTF image does not carry a desired width.</summary>
    public double DefaultImageWidth { get; set; } = 144;

    /// <summary>Default image height in PDF points when the RTF image does not carry a desired height.</summary>
    public double DefaultImageHeight { get; set; } = 96;

    /// <summary>When true, RTF title, author, subject, and keywords are copied to PDF metadata.</summary>
    public bool IncludeMetadata { get; set; } = true;

    /// <summary>When true, RTF tables are converted to PDF tables.</summary>
    public bool IncludeTables { get; set; } = true;

    /// <summary>Returns a normalized copy with valid dimensions and independent PDF options.</summary>
    internal RtfPdfSaveOptions Normalize() {
        if (DefaultImageWidth <= 0) {
            throw new ArgumentOutOfRangeException(nameof(DefaultImageWidth), "Default image width must be greater than zero.");
        }

        if (DefaultImageHeight <= 0) {
            throw new ArgumentOutOfRangeException(nameof(DefaultImageHeight), "Default image height must be greater than zero.");
        }

        return new RtfPdfSaveOptions {
            PdfOptions = PdfOptions?.Clone(),
            IncludeHiddenText = IncludeHiddenText,
            IncludeImages = IncludeImages,
            DefaultImageWidth = DefaultImageWidth,
            DefaultImageHeight = DefaultImageHeight,
            IncludeMetadata = IncludeMetadata,
            IncludeTables = IncludeTables
        };
    }
}
