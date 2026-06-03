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

    /// <summary>Uses a caller-supplied TrueType font family for the default generated text style.</summary>
    public PdfTextStyleCompose FontFamily(PdfEmbeddedFontFamily fontFamily) {
        _opts.UseDefaultTextFontFamily(fontFamily);
        return this;
    }

    /// <summary>Uses caller-supplied TrueType font files for the default generated text style.</summary>
    public PdfTextStyleCompose FontFamily(
        string familyName,
        byte[] regular,
        byte[]? bold = null,
        byte[]? italic = null,
        byte[]? boldItalic = null) {
        _opts.UseDefaultTextFontFamily(new PdfEmbeddedFontFamily(familyName, regular, bold, italic, boldItalic));
        return this;
    }

    /// <summary>Uses caller-supplied TrueType font files for the default generated text style.</summary>
    public PdfTextStyleCompose FontFamily(
        string familyName,
        string regularPath,
        string? boldPath = null,
        string? italicPath = null,
        string? boldItalicPath = null) {
        _opts.UseDefaultTextFontFamily(PdfEmbeddedFontFamily.FromFiles(familyName, regularPath, boldPath, italicPath, boldItalicPath));
        return this;
    }
}

