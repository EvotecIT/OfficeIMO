namespace OfficeIMO.Pdf;

/// <summary>
/// Inline text segment with basic styling.
/// </summary>
public sealed class TextRun {
    public string Text { get; }
    public bool Bold { get; }
    public bool Underline { get; }
    public bool Strike { get; }
    public bool Italic { get; }
    public PdfColor? Color { get; }
    public string? LinkUri { get; }

    public TextRun(string text, bool bold = false, bool underline = false, PdfColor? color = null, bool italic = false, bool strike = false, string? linkUri = null) {
        Text = text ?? string.Empty;
        Bold = bold;
        Underline = underline;
        Italic = italic;
        Strike = strike;
        Color = color;
        LinkUri = linkUri;
    }

    public static TextRun Normal(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: false);
    public static TextRun Bolded(string text, PdfColor? color = null) => new TextRun(text, bold: true, underline: false, color: color, italic: false, strike: false);
    public static TextRun Underlined(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: true, color: color, italic: false, strike: false);
    public static TextRun Italicized(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: false, color: color, italic: true, strike: false);
    public static TextRun BoldUnderlined(string text, PdfColor? color = null) => new TextRun(text, bold: true, underline: true, color: color, italic: false, strike: false);
    public static TextRun BoldItalic(string text, PdfColor? color = null) => new TextRun(text, bold: true, underline: false, color: color, italic: true, strike: false);
    public static TextRun Strikethrough(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: true);
    public static TextRun Link(string text, string uri, PdfColor? color = null, bool underline = true) => new TextRun(text, bold: false, underline: underline, color: color, italic: false, strike: false, linkUri: uri);
}
