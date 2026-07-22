using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

/// <summary>
/// Controls conversion from an RTF document model to a first-party PDF document.
/// </summary>
public sealed class RtfPdfSaveOptions {
    /// <summary>Default number of distinct document font families allowed to probe installed system fonts.</summary>
    public const int DefaultMaximumSystemFontFamilies = 32;

    private PdfCore.PdfResourcePolicy _resourcePolicy = PdfCore.PdfResourcePolicy.CreateDefault();
    /// <summary>Creates RTF to PDF save options.</summary>
    public RtfPdfSaveOptions() {
    }

    /// <summary>Optional PDF engine options. The converter clones the instance before applying RTF page setup.</summary>
    public PdfCore.PdfOptions? PdfOptions { get; set; }

    /// <summary>Host-resource policy. Defaults to balanced conversion: system fonts and bounded in-source resources are allowed, while local and remote reads are denied.</summary>
    public PdfCore.PdfResourcePolicy ResourcePolicy {
        get => _resourcePolicy;
        set => _resourcePolicy = value ?? throw new ArgumentNullException(nameof(value));
    }

    internal PdfCore.PdfConversionReport Report { get; } = new PdfCore.PdfConversionReport();

    /// <summary>When true, RTF hidden runs are included in PDF output. Hidden text is skipped by default.</summary>
    public bool IncludeHiddenText { get; set; }

    /// <summary>When true, supported top-level and inline PNG/JPEG images are emitted to the PDF.</summary>
    public bool IncludeImages { get; set; } = true;

    /// <summary>
    /// Optional converter for image formats that the managed drawing layer cannot rasterize, such as WMF and EMF.
    /// The callback must return a raster payload supported by the shared Drawing-to-PDF pipeline, or null when it cannot convert the image.
    /// </summary>
    public Func<RtfImage, byte[]?>? ImageConverter { get; set; }

    /// <summary>Default image width in PDF points when the RTF image does not carry a desired width.</summary>
    public double DefaultImageWidth { get; set; } = 144;

    /// <summary>Default image height in PDF points when the RTF image does not carry a desired height.</summary>
    public double DefaultImageHeight { get; set; } = 96;

    /// <summary>When true, RTF title, author, subject, and keywords are copied to PDF metadata.</summary>
    public bool IncludeMetadata { get; set; } = true;

    /// <summary>When true, RTF tables are converted to PDF tables.</summary>
    public bool IncludeTables { get; set; } = true;

    /// <summary>When true, semantic RTF header and footer text is mapped to PDF running header/footer text.</summary>
    public bool IncludeHeaderFooters { get; set; } = true;

    /// <summary>When true, footnote, endnote, and annotation bodies referenced by runs are appended to PDF output.</summary>
    public bool IncludeNotes { get; set; } = true;

    /// <summary>
    /// Maximum number of distinct RTF font families allowed to trigger installed-font resolution.
    /// Additional families use dependency-free PDF fallbacks. Set to zero to disable system-font probes.
    /// </summary>
    public int MaximumSystemFontFamilies { get; set; } = DefaultMaximumSystemFontFamilies;

    /// <summary>Returns a normalized copy with valid dimensions and independent PDF options.</summary>
    internal RtfPdfSaveOptions CloneForConversion() {
        if (DefaultImageWidth <= 0) {
            throw new ArgumentOutOfRangeException(nameof(DefaultImageWidth), "Default image width must be greater than zero.");
        }

        if (DefaultImageHeight <= 0) {
            throw new ArgumentOutOfRangeException(nameof(DefaultImageHeight), "Default image height must be greater than zero.");
        }
        if (MaximumSystemFontFamilies < 0) {
            throw new ArgumentOutOfRangeException(nameof(MaximumSystemFontFamilies));
        }

        return new RtfPdfSaveOptions {
            PdfOptions = PdfOptions?.Clone(),
            ResourcePolicy = ResourcePolicy.Clone(),
            IncludeHiddenText = IncludeHiddenText,
            IncludeImages = IncludeImages,
            ImageConverter = ImageConverter,
            DefaultImageWidth = DefaultImageWidth,
            DefaultImageHeight = DefaultImageHeight,
            IncludeMetadata = IncludeMetadata,
            IncludeTables = IncludeTables,
            IncludeHeaderFooters = IncludeHeaderFooters,
            IncludeNotes = IncludeNotes,
            MaximumSystemFontFamilies = MaximumSystemFontFamilies
        };
    }

}
