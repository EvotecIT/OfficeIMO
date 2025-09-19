namespace OfficeIMO.Pdf;

internal sealed class PanelParagraphBlock : IPdfBlock {
    public System.Collections.Generic.List<TextRun> Runs { get; } = new();
    public PdfAlign Align { get; }
    public PdfColor? DefaultColor { get; }
    public PanelStyle Style { get; }
    public PanelParagraphBlock(System.Collections.Generic.IEnumerable<TextRun> runs, PdfAlign align, PdfColor? defaultColor, PanelStyle style) {
        Align = align; DefaultColor = defaultColor; Style = style; Runs.AddRange(runs);
    }
}
