namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent builder for rich paragraphs made of styled text runs.
/// </summary>
public sealed class PdfParagraphBuilder {
    private readonly System.Collections.Generic.List<TextRun> _runs = new();
    private PdfColor? _currentColor;
    private bool _currentBold;
    private bool _currentItalic;
    private bool _currentUnderline;
    private bool _currentStrike;

    /// <summary>Paragraph alignment.</summary>
    public PdfAlign Align { get; }
    /// <summary>Default text color applied when no run color is specified.</summary>
    public PdfColor? DefaultColor { get; }

    /// <summary>Create a new paragraph builder.</summary>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    public PdfParagraphBuilder(PdfAlign align, PdfColor? defaultColor) {
        Align = align; DefaultColor = defaultColor; _currentColor = defaultColor;
    }

    /// <summary>Sets the current run color.</summary>
    public PdfParagraphBuilder Color(PdfColor color) { _currentColor = color; return this; }
    /// <summary>Enables or disables bold for subsequent runs.</summary>
    public PdfParagraphBuilder Bold(bool enable = true) { _currentBold = enable; return this; }
    /// <summary>Enables or disables italic for subsequent runs.</summary>
    public PdfParagraphBuilder Italic(bool enable = true) { _currentItalic = enable; return this; }
    /// <summary>Enables or disables underline for subsequent runs.</summary>
    public PdfParagraphBuilder Underline(bool enable = true) { _currentUnderline = enable; return this; }
    /// <summary>Enables or disables strikethrough for subsequent runs.</summary>
    public PdfParagraphBuilder Strike(bool enable = true) { _currentStrike = enable; return this; }

    /// <summary>Adds a text run using the current style flags.</summary>
    public PdfParagraphBuilder Text(string text) { _runs.Add(new TextRun(text, _currentBold, _currentUnderline, _currentColor, _currentItalic, _currentStrike)); return this; }
    /// <summary>Adds a bold text run.</summary>
    public PdfParagraphBuilder Bold(string text, PdfColor? color = null) { _runs.Add(new TextRun(text, bold: true, underline: false, color: color ?? _currentColor, italic: false)); return this; }
    /// <summary>Adds an italic text run.</summary>
    public PdfParagraphBuilder Italic(string text, PdfColor? color = null) { _runs.Add(new TextRun(text, bold: false, underline: false, color: color ?? _currentColor, italic: true)); return this; }
    /// <summary>Adds an underlined text run.</summary>
    public PdfParagraphBuilder Underlined(string text, PdfColor? color = null) { _runs.Add(new TextRun(text, bold: false, underline: true, color: color ?? _currentColor, italic: false)); return this; }
    /// <summary>Adds a strikethrough text run.</summary>
    public PdfParagraphBuilder Strikethrough(string text, PdfColor? color = null) { _runs.Add(new TextRun(text, bold: false, underline: false, color: color ?? _currentColor, italic: false, strike: true)); return this; }
    /// <summary>Adds a hyperlink text run.</summary>
    /// <param name="text">Link text.</param>
    /// <param name="uri">Absolute URI to open.</param>
    /// <param name="color">Optional link color.</param>
    /// <param name="underline">Whether to underline the link text (default true).</param>
    public PdfParagraphBuilder Link(string text, string uri, PdfColor? color = null, bool underline = true) { _runs.Add(new TextRun(text, bold: false, underline: underline, color: color ?? _currentColor, italic: false, strike: false, linkUri: uri)); return this; }

    internal RichParagraphBlock Build() => new RichParagraphBlock(_runs, Align, DefaultColor);
}
