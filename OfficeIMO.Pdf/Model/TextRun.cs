namespace OfficeIMO.Pdf;

/// <summary>
/// Inline text segment with basic styling.
/// </summary>
public sealed class TextRun {
    /// <summary>Text content of this run.</summary>
    public string Text { get; }
    /// <summary>True when bold style is applied.</summary>
    public bool Bold { get; }
    /// <summary>True when underline is applied.</summary>
    public bool Underline { get; }
    /// <summary>True when strikethrough is applied.</summary>
    public bool Strike { get; }
    /// <summary>True when italic style is applied.</summary>
    public bool Italic { get; }
    /// <summary>Run foreground color (if any).</summary>
    public PdfColor? Color { get; }
    /// <summary>Optional hyperlink URI associated with this run.</summary>
    public string? LinkUri { get; }

    /// <summary>Create a new run with the specified styles.</summary>
    /// <param name="text">Run text.</param>
    /// <param name="bold">Whether to render bold.</param>
    /// <param name="underline">Whether to underline.</param>
    /// <param name="color">Run color or null to use defaults.</param>
    /// <param name="italic">Whether to render italic.</param>
    /// <param name="strike">Whether to render strikethrough.</param>
    /// <param name="linkUri">Optional absolute URI for link annotation.</param>
    public TextRun(string text, bool bold = false, bool underline = false, PdfColor? color = null, bool italic = false, bool strike = false, string? linkUri = null) {
        Text = text ?? string.Empty;
        Bold = bold;
        Underline = underline;
        Italic = italic;
        Strike = strike;
        Color = color;
        LinkUri = linkUri;
    }

    /// <summary>Create a normal (unstyled) run.</summary>
    public static TextRun Normal(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: false);
    /// <summary>Create a bold run.</summary>
    public static TextRun Bolded(string text, PdfColor? color = null) => new TextRun(text, bold: true, underline: false, color: color, italic: false, strike: false);
    /// <summary>Create an underlined run.</summary>
    public static TextRun Underlined(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: true, color: color, italic: false, strike: false);
    /// <summary>Create an italic run.</summary>
    public static TextRun Italicized(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: false, color: color, italic: true, strike: false);
    /// <summary>Create a bold and underlined run.</summary>
    public static TextRun BoldUnderlined(string text, PdfColor? color = null) => new TextRun(text, bold: true, underline: true, color: color, italic: false, strike: false);
    /// <summary>Create a bold and italic run.</summary>
    public static TextRun BoldItalic(string text, PdfColor? color = null) => new TextRun(text, bold: true, underline: false, color: color, italic: true, strike: false);
    /// <summary>Create a strikethrough run.</summary>
    public static TextRun Strikethrough(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: true);
    /// <summary>Create a hyperlink run that points to a URI.</summary>
    /// <param name="text">Link text.</param>
    /// <param name="uri">Absolute URI.</param>
    /// <param name="color">Optional link color.</param>
    /// <param name="underline">Whether to underline the link text.</param>
    public static TextRun Link(string text, string uri, PdfColor? color = null, bool underline = true) => new TextRun(text, bold: false, underline: underline, color: color, italic: false, strike: false, linkUri: uri);
}
