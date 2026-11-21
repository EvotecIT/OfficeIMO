namespace OfficeIMO.OneNote.Pdf;

/// <summary>Controls visual-preserving PDF generated from the native OneNote page canvas.</summary>
public sealed class OneNoteVisualPdfExportOptions {
    private const string VectorPagesMessage = "Visual PDF pages are written as vector content with real text and are no longer rasterized; this setting has no effect.";

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>PDF generation settings, including encryption and compliance requirements.</summary>
    public OfficeIMO.Pdf.PdfCreateOptions PdfOptions { get; set; } = new OfficeIMO.Pdf.PdfCreateOptions();

    /// <summary>Page rendering options shared with image and HTML export.</summary>
    public OneNotePageRenderingOptions PageRendering { get; set; } = new OneNotePageRenderingOptions();

    /// <summary>Not used. Visual PDF pages are vector content.</summary>
    [Obsolete(VectorPagesMessage)]
    public double RasterScale { get; set; } = 2D;

    /// <summary>Not used. Visual PDF pages are vector content.</summary>
    [Obsolete(VectorPagesMessage)]
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
        Title = Title,
        Author = Author,
        Subject = Subject,
        Keywords = Keywords,
        FailOnLoss = FailOnLoss
    };

    internal void Validate() {
        if (PdfOptions == null) throw new InvalidOperationException("PDF options cannot be null.");
        if (PageRendering == null) throw new InvalidOperationException("Page rendering options cannot be null.");
        PageRendering.Validate();
    }
}
