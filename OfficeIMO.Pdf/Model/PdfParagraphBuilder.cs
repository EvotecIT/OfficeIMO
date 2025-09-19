namespace OfficeIMO.Pdf;

public sealed class PdfParagraphBuilder {
    private readonly System.Collections.Generic.List<TextRun> _runs = new();
    private PdfColor? _currentColor;
    private bool _currentBold;
    private bool _currentItalic;
    private bool _currentUnderline;
    private bool _currentStrike;

    public PdfAlign Align { get; }
    public PdfColor? DefaultColor { get; }

    public PdfParagraphBuilder(PdfAlign align, PdfColor? defaultColor) {
        Align = align; DefaultColor = defaultColor; _currentColor = defaultColor;
    }

    public PdfParagraphBuilder Color(PdfColor color) { _currentColor = color; return this; }
    public PdfParagraphBuilder Bold(bool enable = true) { _currentBold = enable; return this; }
    public PdfParagraphBuilder Italic(bool enable = true) { _currentItalic = enable; return this; }
    public PdfParagraphBuilder Underline(bool enable = true) { _currentUnderline = enable; return this; }
    public PdfParagraphBuilder Strike(bool enable = true) { _currentStrike = enable; return this; }

    public PdfParagraphBuilder Text(string text) { _runs.Add(new TextRun(text, _currentBold, _currentUnderline, _currentColor, _currentItalic, _currentStrike)); return this; }
    public PdfParagraphBuilder Bold(string text, PdfColor? color = null) { _runs.Add(new TextRun(text, bold: true, underline: false, color: color ?? _currentColor, italic: false)); return this; }
    public PdfParagraphBuilder Italic(string text, PdfColor? color = null) { _runs.Add(new TextRun(text, bold: false, underline: false, color: color ?? _currentColor, italic: true)); return this; }
    public PdfParagraphBuilder Underlined(string text, PdfColor? color = null) { _runs.Add(new TextRun(text, bold: false, underline: true, color: color ?? _currentColor, italic: false)); return this; }
    public PdfParagraphBuilder Strikethrough(string text, PdfColor? color = null) { _runs.Add(new TextRun(text, bold: false, underline: false, color: color ?? _currentColor, italic: false, strike: true)); return this; }
    public PdfParagraphBuilder Link(string text, string uri, PdfColor? color = null, bool underline = true) { _runs.Add(new TextRun(text, bold: false, underline: underline, color: color ?? _currentColor, italic: false, strike: false, linkUri: uri)); return this; }

    internal RichParagraphBlock Build() => new RichParagraphBlock(_runs, Align, DefaultColor);
}
