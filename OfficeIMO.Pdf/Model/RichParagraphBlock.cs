namespace OfficeIMO.Pdf;

internal sealed class RichParagraphBlock : IPdfBlock {
    public System.Collections.Generic.List<TextRun> Runs { get; } = new();
    public PdfAlign Align { get; }
    public PdfColor? DefaultColor { get; }
    public RichParagraphBlock(System.Collections.Generic.IEnumerable<TextRun> runs, PdfAlign align, PdfColor? defaultColor) {
        Align = align; DefaultColor = defaultColor; Runs.AddRange(runs);
    }
}

