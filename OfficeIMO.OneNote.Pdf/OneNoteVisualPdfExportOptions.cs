namespace OfficeIMO.OneNote.Pdf;

/// <summary>Controls visual-preserving PDF generated from the native OneNote page canvas.</summary>
public sealed class OneNoteVisualPdfExportOptions {
    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>PDF generation settings, including encryption and compliance requirements.</summary>
    public OfficeIMO.Pdf.PdfCreateOptions PdfOptions { get; set; } = new OfficeIMO.Pdf.PdfCreateOptions();

    /// <summary>Page rendering options shared with image and HTML export.</summary>
    public OneNotePageRenderingOptions PageRendering { get; set; } = new OneNotePageRenderingOptions();

    /// <summary>Raster scale used for the full-fidelity page image; 2 equals 144 DPI.</summary>
    public double RasterScale { get; set; } = 2D;

    /// <summary>Maximum decoded pixels allocated for one PDF page image.</summary>
    public long MaximumRasterPixels { get; set; } = 100_000_000L;

    /// <summary>Optional PDF title metadata. The notebook or section name is used when absent.</summary>
    public string? Title { get; set; }

    /// <summary>Optional PDF author metadata.</summary>
    public string? Author { get; set; }

    /// <summary>Optional PDF subject metadata.</summary>
    public string? Subject { get; set; }

    /// <summary>Optional PDF keywords metadata.</summary>
    public string? Keywords { get; set; }

    /// <summary>Creates an independent copy of these settings.</summary>
    public OneNoteVisualPdfExportOptions Clone() => new OneNoteVisualPdfExportOptions {
        PdfOptions = PdfOptions?.Clone() ?? throw new InvalidOperationException("PDF options cannot be null."),
        PageRendering = PageRendering?.Clone() ?? new OneNotePageRenderingOptions(),
        RasterScale = RasterScale,
        MaximumRasterPixels = MaximumRasterPixels,
        Title = Title,
        Author = Author,
        Subject = Subject,
        Keywords = Keywords,
        FailOnLoss = FailOnLoss
    };

    internal void Validate() {
        if (PdfOptions == null) throw new InvalidOperationException("PDF options cannot be null.");
        if (PageRendering == null) throw new InvalidOperationException("Page rendering options cannot be null.");
        if (double.IsNaN(RasterScale) || double.IsInfinity(RasterScale) || RasterScale <= 0D) throw new ArgumentOutOfRangeException(nameof(RasterScale));
        if (MaximumRasterPixels < 1L) throw new ArgumentOutOfRangeException(nameof(MaximumRasterPixels));
    }
}
