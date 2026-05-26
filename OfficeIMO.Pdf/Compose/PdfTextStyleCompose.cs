namespace OfficeIMO.Pdf;

/// <summary>Builder for default text styling applied to page content.</summary>
public class PdfTextStyleCompose {
    private readonly PdfOptions _opts;
    internal PdfTextStyleCompose(PdfOptions opts) { _opts = opts; }
    /// <summary>Sets the default font size (points).</summary>
    public PdfTextStyleCompose FontSize(double size) { Guard.Positive(size, nameof(size)); _opts.DefaultFontSize = size; return this; }
    /// <summary>Sets the default text color.</summary>
    public PdfTextStyleCompose Color(PdfColor color) { _opts.DefaultTextColor = color; return this; }
    /// <summary>Sets the default standard font family.</summary>
    public PdfTextStyleCompose Font(PdfStandardFont font) {
        Guard.StandardFont(font, nameof(font), "PDF default font must be one of the supported standard PDF fonts.");
        _opts.DefaultFont = font;
        return this;
    }
}

